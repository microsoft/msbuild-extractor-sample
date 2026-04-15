using Microsoft.Build.Logging.StructuredLogger;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MSBuild.CompileCommands.Extractor
{
    public class OutOfProcessExtractor : ICompileCommandsExtractor
    {
        private readonly string _msbuildPath;
        private readonly string _projectPath;
        private readonly string _configuration;
        private readonly string _platform;
        private readonly bool _enableLogger;
        private readonly string? _solutionDir;
        private readonly string? _vcToolsInstallDir;
        private readonly string? _vcTargetsPath;

        public OutOfProcessExtractor(
            string msbuildPath,
            string projectPath,
            string configuration = "Debug",
            string platform = "x64",
            bool enableLogger = false,
            string? solutionDir = null,
            string? vcToolsInstallDir = null,
            string? vcTargetsPath = null)
        {
            _msbuildPath = msbuildPath;
            _projectPath = Path.GetFullPath(projectPath);
            _configuration = configuration;
            _platform = platform;
            _enableLogger = enableLogger;
            _solutionDir = solutionDir;
            _vcToolsInstallDir = vcToolsInstallDir?.TrimEnd('\\');
            if (_vcToolsInstallDir != null && !_vcToolsInstallDir.StartsWith(@"\\"))
                _vcToolsInstallDir = _vcToolsInstallDir.Replace(@"\\", @"\");
            _vcTargetsPath = vcTargetsPath;

            if (!File.Exists(_msbuildPath))
                throw new FileNotFoundException($"MSBuild not found: {_msbuildPath}");
            if (_msbuildPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
            {
                // MSBuild.dll from a .NET SDK can't be launched directly — use dotnet exec
                var dotnetPath = FindDotnetExe();
                if (dotnetPath == null)
                    throw new FileNotFoundException(
                        $"MSBuild path '{_msbuildPath}' is a .dll and requires 'dotnet' to execute, but dotnet was not found on PATH.");
                _dotnetExePath = dotnetPath;
            }
            if (!File.Exists(_projectPath))
                throw new FileNotFoundException($"Project file not found: {_projectPath}");
        }

        private readonly string? _dotnetExePath;

        private string[] _externalIncludePaths = [];
        private string[] _includePaths = [];

        private static string? TryResolveMsBuildExpression(string expr)
        {
            var match = Regex.Match(expr,
                @"\$\(\[MSBuild\]::NormalizePath\('([^']+)'(?:.*?)(?:\)\))?(.*)$");
            if (match.Success)
            {
                var basePath = match.Groups[1].Value;
                var suffix = match.Groups[2].Value.TrimStart('\\', '\'', ',', ' ');
                try
                {
                    var resolved = Path.GetFullPath(Path.Combine(basePath, suffix));
                    if (Directory.Exists(resolved))
                        return resolved;
                }
                catch { }
            }
            return null;
        }

        private static IEnumerable<string> ResolveIncludePathEntries(IEnumerable<string> entries)
        {
            foreach (var p in entries)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                if (!p.StartsWith("$("))
                {
                    string resolved;
                    try { resolved = Uri.UnescapeDataString(p.Trim()); } catch { resolved = p.Trim(); }
                    if (!string.IsNullOrEmpty(resolved))
                        yield return resolved;
                }
                else
                {
                    var resolved = TryResolveMsBuildExpression(p);
                    if (resolved != null)
                        yield return resolved;
                }
            }
        }

        private static IEnumerable<string> GetVsAuxiliaryIncludePaths(string? vcToolsInstallDir)
        {
            if (string.IsNullOrEmpty(vcToolsInstallDir)) return [];
            try
            {
                var vcDir = Path.GetFullPath(Path.Combine(vcToolsInstallDir, "..", "..", ".."));
                var results = new List<string>();
                var auxInclude = Path.Combine(vcDir, "Auxiliary", "VS", "include");
                if (Directory.Exists(auxInclude))
                    results.Add(auxInclude);
                var auxUnitTest = Path.Combine(vcDir, "Auxiliary", "VS", "UnitTest", "include");
                if (Directory.Exists(auxUnitTest))
                    results.Add(auxUnitTest);
                return results;
            }
            catch { return []; }
        }

        private static string? FindDotnetExe()
        {
            var pathVar = Environment.GetEnvironmentVariable("PATH") ?? "";
            var exeName = OperatingSystem.IsWindows() ? "dotnet.exe" : "dotnet";
            foreach (var dir in pathVar.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
            {
                var candidate = Path.Combine(dir.Trim(), exeName);
                if (File.Exists(candidate))
                    return candidate;
            }
            return null;
        }

        public List<CompileCommand> ExtractCompileCommands()
        {
            var entries = GetClCommandLines();
            var commands = entries.Count > 0 ? ToCompileCommands(entries) : new List<CompileCommand>();
            if (commands.Count > 0)
                return commands;

            // Fallback for GN-generated projects (Chromium, Crashpad, WebRTC)
            var fallback = TryExtractFromItemDefinitionGroup();
            if (fallback.Count > 0)
            {
                Console.Error.WriteLine($"Info: Used ItemDefinitionGroup fallback for {_projectPath} ({fallback.Count} entries)");
                return fallback;
            }

            Console.Error.WriteLine($"Warning: No compile commands found in {_projectPath}");
            return new List<CompileCommand>();
        }

        public void WriteCompileCommandsJson(string? outputPath = null)
        {
            var commands = ExtractCompileCommands();
            outputPath ??= Path.Combine(Path.GetDirectoryName(_projectPath)!, "compile_commands.json");

            var json = JsonSerializer.Serialize(commands, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            File.WriteAllText(outputPath, json);
        }

        private List<ClCommandLineEntry> GetClCommandLines()
        {
            var binlogPath = Path.Combine(Path.GetTempPath(), $"msbuild_{Guid.NewGuid():N}.binlog");

            try
            {
                RunMSBuild(binlogPath);
                return ParseBinlog(binlogPath);
            }
            catch (Exception ex)
            {
                // GN-generated projects may fail the design-time build.
                // Return empty to trigger the ItemDefinitionGroup fallback.
                if (_enableLogger)
                    Console.Error.WriteLine($"Design-time build failed for {_projectPath}: {ex.Message} (will try fallback)");
                return new List<ClCommandLineEntry>();
            }
            finally
            {
                try { if (File.Exists(binlogPath) && !_enableLogger) File.Delete(binlogPath); }
                catch { /* cleanup failure is not fatal */ }
            }
        }

        private void RunMSBuild(string binlogPath)
        {
            var properties = new List<KeyValuePair<string, string>>
            {
                new("Configuration", _configuration),
                new("Platform", _platform),
                new("DesignTimeBuild", "true"),
                new("BuildingInsideVisualStudio", "true"),
                new("BuildProjectReferences", "true")
            };

            if (_solutionDir != null)
            {
                var dir = _solutionDir.EndsWith('\\') ? _solutionDir : _solutionDir + "\\";
                properties.Add(new("SolutionDir", dir));
            }

            if (_vcToolsInstallDir != null)
            {
                var dir = _vcToolsInstallDir.StartsWith(@"\\")
                    ? _vcToolsInstallDir
                    : _vcToolsInstallDir.Replace(@"\\", @"\");
                dir = dir.EndsWith('\\') ? dir : dir + "\\";
                properties.Add(new("VCToolsInstallDir", dir));
            }

            // Don't set WindowsTargetPlatformVersion as a global property — it would
            // override projects with explicit full versions (e.g. 10.0.19041.0).
            // Instead, set as environment variable which has lower precedence in MSBuild.
            var latestSdk = VsWhereHelper.FindLatestWindowsSdkVersion();
            if (latestSdk != null)
            {
                Environment.SetEnvironmentVariable("WindowsTargetPlatformVersion", latestSdk);
            }

            // Use ArgumentList to avoid shell quoting issues
            // with paths containing spaces and trailing backslashes
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            if (_dotnetExePath != null)
            {
                // MSBuild.dll: run via "dotnet exec MSBuild.dll ..."
                startInfo.FileName = _dotnetExePath;
                startInfo.ArgumentList.Add("exec");
                startInfo.ArgumentList.Add(_msbuildPath);
            }
            else
            {
                startInfo.FileName = _msbuildPath;
            }

            startInfo.ArgumentList.Add(_projectPath);
            startInfo.ArgumentList.Add("/t:GetClCommandLines");

            // Use separate /p: per property to avoid issues with semicolons
            // or special characters in property values
            foreach (var prop in properties)
                startInfo.ArgumentList.Add($"/p:{prop.Key}={prop.Value}");

            startInfo.ArgumentList.Add($"/bl:{binlogPath}");
            startInfo.ArgumentList.Add("/nologo");
            startInfo.ArgumentList.Add("/v:quiet");

            if (_vcTargetsPath != null)
            {
                var path = _vcTargetsPath.EndsWith('\\') ? _vcTargetsPath : _vcTargetsPath + "\\";
                startInfo.Environment["VCTargetsPath"] = path;
            }

            if (_vcToolsInstallDir != null)
            {
                var dir = _vcToolsInstallDir.StartsWith(@"\\")
                    ? _vcToolsInstallDir
                    : _vcToolsInstallDir.Replace(@"\\", @"\");
                dir = dir.EndsWith('\\') ? dir : dir + "\\";
                startInfo.Environment["VCToolsInstallDir"] = dir;
            }

            if (_enableLogger)
            {
                var cmdLine = _dotnetExePath != null
                    ? $"dotnet exec \"{_msbuildPath}\""
                    : $"\"{_msbuildPath}\"";
                var args = string.Join(" ", startInfo.ArgumentList
                    .Skip(_dotnetExePath != null ? 2 : 0)
                    .Select(a => a.Contains(' ') ? $"\"{a}\"" : a));
                Console.WriteLine($"Running: {cmdLine} {args}");
            }

            using var process = System.Diagnostics.Process.Start(startInfo);
            if (process == null)
                throw new Exception("Failed to start MSBuild process");

            // Read both streams as tasks to avoid deadlock when both buffers fill
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            // Enforce timeout FIRST — ReadToEnd blocks forever if the process hangs
            if (!process.WaitForExit(300000))
            {
                try { process.Kill(true); } catch { }
                throw new Exception("MSBuild timed out after 5 minutes");
            }

            // Now safe to read (process has exited)
            var output = stdoutTask.Result;
            var error = errorTask.Result;

            if (_enableLogger)
            {
                if (!string.IsNullOrWhiteSpace(output)) Console.WriteLine(output);
                if (!string.IsNullOrWhiteSpace(error)) Console.Error.WriteLine(error);
                Console.WriteLine($"Binlog: {binlogPath}");
            }

            if (process.ExitCode != 0)
                throw new Exception($"MSBuild failed with exit code {process.ExitCode}:\n{error}\n{output}");
        }

        private List<ClCommandLineEntry> ParseBinlog(string binlogPath)
        {
            var entries = new List<ClCommandLineEntry>();

            var build = BinaryLog.ReadBuild(binlogPath);
            var target = build.FindFirstDescendant<Target>(t => t.Name == "GetClCommandLines");

            // Extract ExternalIncludePath property from binlog
            var externalIncludeProp = build.FindChildrenRecursive<Property>()
                .LastOrDefault(p => p.Name == "ExternalIncludePath" && !string.IsNullOrEmpty(p.Value));
            if (externalIncludeProp != null)
            {
                _externalIncludePaths = ResolveIncludePathEntries(
                    externalIncludeProp.Value.Split(';', StringSplitOptions.RemoveEmptyEntries)).ToArray();
            }

            // Extract IncludePath property from binlog. This property feeds the
            // compiler's system include directories. Projects like SDL add source
            // directories here instead of AdditionalIncludeDirectories, and
            // GetClCommandLines does not emit /I flags for them.
            var includePathProp = build.FindChildrenRecursive<Property>()
                .LastOrDefault(p => p.Name == "IncludePath" && !string.IsNullOrEmpty(p.Value));
            if (includePathProp != null)
            {
                var externalSet = new HashSet<string>(
                    _externalIncludePaths, StringComparer.OrdinalIgnoreCase);
                _includePaths = ResolveIncludePathEntries(
                    includePathProp.Value.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    .Where(p => !externalSet.Contains(p))
                    .ToArray();
            }

            // Add VS auxiliary include paths from VCToolsInstallDir
            var vcToolsDirProp = build.FindChildrenRecursive<Property>()
                .LastOrDefault(p => p.Name == "VCToolsInstallDir" && !string.IsNullOrEmpty(p.Value));
            var vcToolsDir = _vcToolsInstallDir ?? vcToolsDirProp?.Value?.TrimEnd('\\');
            var allIncludes = new HashSet<string>(_externalIncludePaths.Concat(_includePaths), StringComparer.OrdinalIgnoreCase);
            var auxPaths = GetVsAuxiliaryIncludePaths(vcToolsDir)
                .Where(p => !allIncludes.Contains(p))
                .ToArray();
            if (auxPaths.Length > 0)
                _includePaths = _includePaths.Concat(auxPaths).ToArray();

            if (target == null)
            {
                if (_enableLogger)
                    Console.WriteLine("Warning: GetClCommandLines target not found in binlog");
                return entries;
            }

            // Try AddItem nodes first (older binlog format)
            foreach (var child in target.Children)
            {
                if (child is AddItem addItem)
                {
                    var entry = ParseAddItem(addItem);
                    if (entry != null) entries.Add(entry);
                }
            }

            // If no AddItem nodes found, try Item nodes in TargetOutputs folder (newer binlog format)
            if (entries.Count == 0)
            {
                var outputFolder = target.FindFirstDescendant<Folder>(f => f.Name == "TargetOutputs");
                if (outputFolder != null)
                {
                    foreach (var child in outputFolder.Children)
                    {
                        if (child is Item item)
                        {
                            var entry = ParseItem(item);
                            if (entry != null) entries.Add(entry);
                        }
                    }
                }
            }

            return entries;
        }

        private ClCommandLineEntry? ParseAddItem(AddItem addItem)
        {
            var commandLine = addItem.Name ?? "";
            var workingDir = "";
            var toolPath = "";
            var files = new List<string>();

            foreach (var metadata in addItem.Children.OfType<Metadata>())
            {
                switch (metadata.Name)
                {
                    case "WorkingDirectory": workingDir = metadata.Value; break;
                    case "ToolPath": toolPath = metadata.Value; break;
                    case "Files":
                        files.AddRange(metadata.Value.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(f => f.Trim()));
                        break;
                }
            }

            return files.Count > 0
                ? new ClCommandLineEntry(commandLine, workingDir, toolPath, files.ToArray())
                : null;
        }

        private ClCommandLineEntry? ParseItem(Item item)
        {
            var commandLine = item.Name ?? "";
            var workingDir = "";
            var toolPath = "";
            var files = new List<string>();

            foreach (var metadata in item.Children.OfType<Metadata>())
            {
                switch (metadata.Name)
                {
                    case "WorkingDirectory": workingDir = metadata.Value; break;
                    case "ToolPath": toolPath = metadata.Value; break;
                    case "Files":
                        files.AddRange(metadata.Value.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(f => f.Trim()));
                        break;
                }
            }

            return files.Count > 0
                ? new ClCommandLineEntry(commandLine, workingDir, toolPath, files.ToArray())
                : null;
        }

        private List<CompileCommand> ToCompileCommands(List<ClCommandLineEntry> entries)
        {
            var commands = new List<CompileCommand>();
            var projectName = Path.GetFileNameWithoutExtension(_projectPath);

            foreach (var entry in entries)
            {
                var toolPath = ResolveToolPath(entry.ToolPath);
                var baseArgs = new List<string> { toolPath };

                // Inject --target so clangd uses the correct architecture triple.
                var target = GetClangTargetTriple();
                if (target != null)
                    baseArgs.Add($"--target={target}");

                // Disable clang's error limit to prevent cascading failures.
                // MSVC STL headers can trigger errors that clang recovers from,
                // but the default limit (20) causes fatal_too_many_errors which
                // stops parsing and cascades into hundreds of false diagnostics.
                baseArgs.Add("-ferror-limit=0");

                foreach (var includePath in _externalIncludePaths)
                {
                    baseArgs.Add($"/I{includePath}");
                }

                foreach (var includePath in _includePaths)
                {
                    baseArgs.Add($"/I{includePath}");
                }

                baseArgs.AddRange(CommandLineTokenizer.TokenizeWithResponseFiles(
                    entry.CommandLine, entry.WorkingDirectory));

                // Merge two-arg /external:I flags into single token.
                // MSBuild emits "/external:I" "<path>" as separate tokens but
                // cl.exe also accepts the concatenated form "/external:I<path>".
                var mergedArgs = new List<string>();
                for (int i = 0; i < baseArgs.Count; i++)
                {
                    if (i + 1 < baseArgs.Count &&
                        baseArgs[i].Equals("/external:I", StringComparison.OrdinalIgnoreCase))
                    {
                        mergedArgs.Add($"/external:I{baseArgs[i + 1]}");
                        i++; // skip next
                    }
                    else
                    {
                        mergedArgs.Add(baseArgs[i]);
                    }
                }
                baseArgs = mergedArgs;

                // Strip build-output flags irrelevant for IntelliSense
                baseArgs.RemoveAll(a =>
                    a.StartsWith("/Fo", StringComparison.OrdinalIgnoreCase) ||
                    a.StartsWith("/Fd", StringComparison.OrdinalIgnoreCase) ||
                    a.StartsWith("/errorReport", StringComparison.OrdinalIgnoreCase));

                foreach (var file in entry.Files)
                {
                    var fileArgs = new List<string>(baseArgs) { file };

                    commands.Add(new CompileCommand(
                        File: file,
                        Arguments: fileArgs.ToArray(),
                        Directory: entry.WorkingDirectory,
                        ProjectPath: _projectPath,
                        ProjectName: projectName,
                        Configuration: _configuration,
                        Platform: _platform
                    ));
                }
            }

            return commands;
        }

        private string ResolveToolPath(string toolPath)
        {
            if (toolPath.Contains("system32", StringComparison.OrdinalIgnoreCase) &&
                toolPath.EndsWith("CL.exe", StringComparison.OrdinalIgnoreCase))
            {
                var resolved = VsWhereHelper.ResolveClExePath(_vcToolsInstallDir, _platform);
                if (resolved != null)
                    return resolved;
            }
            return toolPath;
        }

        /// <summary>
        /// Maps the MSBuild Platform to a clang target triple so that clangd
        /// compiles with the correct pointer size, predefined macros (_WIN64,
        /// _M_X64, etc.), and ABI.
        /// </summary>
        private string? GetClangTargetTriple()
        {
            return _platform.ToLowerInvariant() switch
            {
                "win32" or "x86" => "i686-pc-windows-msvc",
                "x64" or "x86_64" or "amd64" => "x86_64-pc-windows-msvc",
                "arm" => "thumbv7-pc-windows-msvc",
                "arm64" or "aarch64" => "aarch64-pc-windows-msvc",
                _ => null
            };
        }

        private static readonly HashSet<string> CppExtensions = new(StringComparer.OrdinalIgnoreCase)
            { ".cpp", ".cc", ".cxx", ".c", ".c++", ".cp" };

        /// <summary>
        /// Scans disk for C/C++ source files when the project XML contains none.
        /// Uses header references as directory hints, falling back to the project directory.
        /// </summary>
        private static List<string> FindSourceFilesOnDisk(string projectDir, XDocument doc, XNamespace ns)
        {
            var files = new List<string>();

            // Collect directories from ClInclude and None items (header references)
            var headerPaths = doc.Root?
                .Elements(ns + "ItemGroup")
                .SelectMany(ig => ig.Elements(ns + "ClInclude").Concat(ig.Elements(ns + "None")))
                .Select(e => e.Attribute("Include")?.Value)
                .Where(v => v != null)
                .Cast<string>()
                .Select(v => Path.GetDirectoryName(Path.GetFullPath(Path.Combine(projectDir, v))))
                .Where(d => d != null && Directory.Exists(d))
                .Cast<string>()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            if (headerPaths.Count > 0)
            {
                foreach (var dir in headerPaths)
                {
                    try
                    {
                        files.AddRange(Directory.EnumerateFiles(dir!, "*.*", SearchOption.TopDirectoryOnly)
                            .Where(f => CppExtensions.Contains(Path.GetExtension(f))));
                    }
                    catch (Exception) { /* permission denied, etc. */ }
                }
            }
            else
            {
                try
                {
                    files.AddRange(Directory.EnumerateFiles(projectDir, "*.*", SearchOption.TopDirectoryOnly)
                        .Where(f => CppExtensions.Contains(Path.GetExtension(f))));
                }
                catch (Exception) { /* permission denied, etc. */ }
            }

            return files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>
        /// Fallback extraction for GN-generated projects (Chromium, Crashpad, WebRTC).
        /// Parses the vcxproj XML directly to read ItemDefinitionGroup/ClCompile settings
        /// and CustomBuild source file items.
        /// </summary>
        private List<CompileCommand> TryExtractFromItemDefinitionGroup()
        {
            XDocument doc;
            try
            {
                doc = XDocument.Load(_projectPath);
            }
            catch
            {
                return new List<CompileCommand>();
            }

            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var projectDir = Path.GetDirectoryName(_projectPath)!;
            var projectName = Path.GetFileNameWithoutExtension(_projectPath);

            // Find ClCompile settings in ItemDefinitionGroup
            var clCompileDefElement = doc.Root?
                .Elements(ns + "ItemDefinitionGroup")
                .Elements(ns + "ClCompile")
                .FirstOrDefault();

            if (clCompileDefElement == null)
                return new List<CompileCommand>();

            string GetElementValue(string name) =>
                clCompileDefElement.Element(ns + name)?.Value ?? "";

            var additionalIncludes = GetElementValue("AdditionalIncludeDirectories");
            var preprocessorDefs = GetElementValue("PreprocessorDefinitions");
            var additionalOptions = GetElementValue("AdditionalOptions");
            var languageStandard = GetElementValue("LanguageStandard");
            var disabledWarnings = GetElementValue("DisableSpecificWarnings");

            // Collect source files from CustomBuild items (GN pattern) or ClCompile items
            var sourceFiles = doc.Root?
                .Elements(ns + "ItemGroup")
                .Elements(ns + "CustomBuild")
                .Select(e => e.Attribute("Include")?.Value)
                .Where(v => v != null && CppExtensions.Contains(Path.GetExtension(v)))
                .Cast<string>()
                .ToList() ?? new List<string>();

            if (sourceFiles.Count == 0)
            {
                sourceFiles = doc.Root?
                    .Elements(ns + "ItemGroup")
                    .Elements(ns + "ClCompile")
                    .Select(e => e.Attribute("Include")?.Value)
                    .Where(v => v != null)
                    .Cast<string>()
                    .ToList() ?? new List<string>();
            }

            // Fallback: scan disk for source files when none found in project XML.
            // GN-generated projects sometimes have empty ItemGroups with source
            // files present on disk nearby.
            if (sourceFiles.Count == 0)
            {
                sourceFiles = FindSourceFilesOnDisk(projectDir, doc, ns);
            }

            if (sourceFiles.Count == 0)
                return new List<CompileCommand>();

            // Build the compiler arguments
            var toolPath = ResolveToolPath(VsWhereHelper.ResolveClExePath(_vcToolsInstallDir, _platform)
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "CL.exe"));

            var baseArgs = new List<string> { toolPath };

            var target = GetClangTargetTriple();
            if (target != null)
                baseArgs.Add($"--target={target}");
            baseArgs.Add("-ferror-limit=0");

            foreach (var p in _externalIncludePaths)
                baseArgs.Add($"/I{p}");
            foreach (var p in _includePaths)
                baseArgs.Add($"/I{p}");

            // Add /I from AdditionalIncludeDirectories
            foreach (var inc in additionalIncludes.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = inc.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("%(") || trimmed.StartsWith("$(")) continue;
                var resolved = Path.IsPathRooted(trimmed)
                    ? trimmed
                    : Path.GetFullPath(Path.Combine(projectDir, trimmed));
                baseArgs.Add($"/I{resolved}");
            }

            // Add /D from PreprocessorDefinitions
            foreach (var def in preprocessorDefs.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = def.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("%(")) continue;
                baseArgs.Add($"/D{trimmed}");
            }

            // Add disabled warnings
            foreach (var w in disabledWarnings.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = w.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("%(")) continue;
                baseArgs.Add($"/wd{trimmed}");
            }

            // Add language standard
            if (!string.IsNullOrEmpty(languageStandard) && !languageStandard.StartsWith("%("))
            {
                var stdFlag = languageStandard.ToLowerInvariant() switch
                {
                    "stdcpp14" => "/std:c++14",
                    "stdcpp17" => "/std:c++17",
                    "stdcpp20" => "/std:c++20",
                    "stdcpplatest" => "/std:c++latest",
                    "stdc11" => "/std:c11",
                    "stdc17" => "/std:c17",
                    _ => null
                };
                if (stdFlag != null)
                    baseArgs.Add(stdFlag);
            }

            // Add AdditionalOptions
            if (!string.IsNullOrEmpty(additionalOptions))
            {
                var tokens = CommandLineTokenizer.TokenizeWithResponseFiles(
                    additionalOptions, projectDir);
                baseArgs.AddRange(tokens.Where(t => !t.StartsWith("%(")));
            }

            // Merge two-arg /external:I flags into single token
            var mergedArgs2 = new List<string>();
            for (int i = 0; i < baseArgs.Count; i++)
            {
                if (i + 1 < baseArgs.Count &&
                    baseArgs[i].Equals("/external:I", StringComparison.OrdinalIgnoreCase))
                {
                    mergedArgs2.Add($"/external:I{baseArgs[i + 1]}");
                    i++;
                }
                else
                {
                    mergedArgs2.Add(baseArgs[i]);
                }
            }
            baseArgs = mergedArgs2;

            // Strip build-output flags
            baseArgs.RemoveAll(a =>
                a.StartsWith("/Fo", StringComparison.OrdinalIgnoreCase) ||
                a.StartsWith("/Fd", StringComparison.OrdinalIgnoreCase) ||
                a.StartsWith("/errorReport", StringComparison.OrdinalIgnoreCase));

            var commands = new List<CompileCommand>();
            foreach (var file in sourceFiles)
            {
                var fullPath = Path.IsPathRooted(file)
                    ? file
                    : Path.GetFullPath(Path.Combine(projectDir, file));

                var fileArgs = new List<string>(baseArgs) { fullPath };

                commands.Add(new CompileCommand(
                    File: fullPath,
                    Arguments: fileArgs.ToArray(),
                    Directory: projectDir,
                    ProjectPath: _projectPath,
                    ProjectName: projectName,
                    Configuration: _configuration,
                    Platform: _platform
                ));
            }

            return commands;
        }

        public static List<CompileCommand> ExtractCompileCommandsFromSolution(
            string msbuildPath, string solutionPath, string configuration = "Debug",
            string platform = "x64", bool enableLogger = false,
            string? vcToolsInstallDir = null, string? vcTargetsPath = null)
        {
            var solutionDir = Path.GetDirectoryName(Path.GetFullPath(solutionPath))!;
            var projects = ProjectDiscovery.GetVcProjectsFromSolution(solutionPath);
            var allCommands = new List<CompileCommand>();

            foreach (var project in projects)
            {
                try
                {
                    var extractor = new OutOfProcessExtractor(
                        msbuildPath, project.Path, configuration, platform, enableLogger,
                        solutionDir, vcToolsInstallDir, vcTargetsPath);
                    allCommands.AddRange(extractor.ExtractCompileCommands());
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Warning: Failed to extract compile commands from {project.Path}: {ex.Message}");
                }
            }

            return allCommands;
        }

        public static void WriteCompileCommandsJsonFromSolution(
            string msbuildPath, string solutionPath, string configuration = "Debug",
            string platform = "x64", bool enableLogger = false, string? outputPath = null,
            string? vcToolsInstallDir = null, string? vcTargetsPath = null)
        {
            var commands = ExtractCompileCommandsFromSolution(msbuildPath, solutionPath, configuration, platform, enableLogger, vcToolsInstallDir, vcTargetsPath);
            outputPath ??= Path.Combine(Path.GetDirectoryName(Path.GetFullPath(solutionPath))!, "compile_commands.json");

            var json = JsonSerializer.Serialize(commands, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            File.WriteAllText(outputPath, json);
        }
    }
}
