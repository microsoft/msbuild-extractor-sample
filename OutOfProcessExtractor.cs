using Microsoft.Build.Logging.StructuredLogger;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MSBuild.CompileCommands.Extractor
{
    public enum MsBuildLauncher { Auto, Cmd, Direct, Dotnet }
    public enum IncludePathOrder { Auto, Prepend, Append }

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
        private readonly string? _clPath;
        private readonly IReadOnlyDictionary<string, string> _userProperties;
        private readonly IReadOnlyDictionary<string, string> _userEnv;
        private readonly MsBuildLauncher _launcher;
        private readonly IncludePathOrder _includePathOrder;
        private readonly bool _emitDefaults;
        private readonly bool _mergeDefaults;

        public OutOfProcessExtractor(
            string msbuildPath,
            string projectPath,
            string configuration = "Debug",
            string platform = "x64",
            bool enableLogger = false,
            string? solutionDir = null,
            string? vcToolsInstallDir = null,
            string? vcTargetsPath = null,
            string? clPath = null,
            IReadOnlyDictionary<string, string>? msbuildProperties = null,
            IReadOnlyDictionary<string, string>? msbuildEnv = null,
            MsBuildLauncher launcher = MsBuildLauncher.Auto,
            IncludePathOrder includePathOrder = IncludePathOrder.Auto,
            bool emitDefaults = false,
            bool mergeDefaults = false)
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
            _clPath = clPath;
            _userProperties = msbuildProperties ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _userEnv = msbuildEnv ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _launcher = launcher;
            _includePathOrder = includePathOrder;
            _emitDefaults = emitDefaults;
            _mergeDefaults = mergeDefaults;

            if (!File.Exists(_msbuildPath))
                throw new FileNotFoundException($"MSBuild not found: {_msbuildPath}");
            if (_msbuildPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
            {
                // MSBuild.dll from a .NET SDK can't be launched directly, use dotnet exec
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

        private static readonly char[] CmdMetaChars =
            { ' ', '\t', ';', '&', '|', '^', '<', '>', '(', ')' };

        /// <summary>
        /// Quotes a command-line argument for cmd.exe. Handles whitespace, cmd
        /// metacharacters (; &amp; | ^ &lt; &gt; ( )), and trailing backslashes
        /// that would otherwise escape the closing quote.
        /// </summary>
        private static string QuoteCmdArg(string arg)
        {
            if (arg.StartsWith('"') && arg.EndsWith('"')) return arg;
            if (arg.IndexOfAny(CmdMetaChars) < 0) return arg;
            int trailingBackslashes = 0;
            for (int i = arg.Length - 1; i >= 0 && arg[i] == '\\'; i--) trailingBackslashes++;
            return "\"" + arg + new string('\\', trailingBackslashes) + "\"";
        }

        private static string? TryResolveMsBuildExpression(string expr)
        {
            var match = Regex.Match(expr, @"\$\(\[MSBuild\]::NormalizePath\('([^']+)'(?:.*?)(?:\)\))?(.*)$");
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
            foreach (var entry in entries)
            {
                if (string.IsNullOrWhiteSpace(entry)) continue;
                if (!entry.StartsWith("$("))
                {
                    string resolved;
                    try { resolved = Uri.UnescapeDataString(entry.Trim()); } catch { resolved = entry.Trim(); }
                    if (!string.IsNullOrEmpty(resolved))
                        yield return resolved;
                }
                else
                {
                    var resolved = TryResolveMsBuildExpression(entry);
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
                // Design-time builds should not build referenced projects.
                new("BuildProjectReferences", "false"),
                // Stable locale for MSBuild error strings.
                new("LangID", "1033")
            };

            // Pass MicrosoftBuildCppTasksCommonPath if VCTargetsPath is set (needed for custom MSBuild environments)
            if (_vcTargetsPath != null)
            {
                properties.Add(new("MicrosoftBuildCppTasksCommonPath", _vcTargetsPath.EndsWith('\\') ? _vcTargetsPath : _vcTargetsPath + "\\"));
            }

            if (_solutionDir != null)
            {
                var dir = _solutionDir.EndsWith('\\') ? _solutionDir : _solutionDir + "\\";
                properties.Add(new("SolutionDir", dir));
            }

            // Skip VCToolsInstallDir when MicrosoftBuildCppTasksCommonPath is set
            // (custom MSBuild environments manage their own tool paths)
            if (_vcToolsInstallDir != null && !properties.Any(p => p.Key == "MicrosoftBuildCppTasksCommonPath"))
            {
                var dir = _vcToolsInstallDir.StartsWith(@"\\")
                    ? _vcToolsInstallDir
                    : _vcToolsInstallDir.Replace(@"\\", @"\");
                dir = dir.EndsWith('\\') ? dir : dir + "\\";
                properties.Add(new("VCToolsInstallDir", dir));

                // Derive VCInstallDir from VCToolsInstallDir (3 levels up)
                // to prevent unresolved fallback values in repo props files.
                try
                {
                    var vcInstallDir = Path.GetFullPath(Path.Combine(dir, "..", "..", ".."));
                    if (Directory.Exists(vcInstallDir))
                    {
                        var vcNorm = vcInstallDir.EndsWith('\\') ? vcInstallDir : vcInstallDir + "\\";
                        properties.Add(new("VCInstallDir", vcNorm));
                    }
                }
                catch { }
            }

            // Detect and set NETFXKitsDir to prevent unresolved fallback values
            var netfxKitsDir = InProcessExtractor.FindNETFXKitsDir();
            if (netfxKitsDir != null && !properties.Any(p => p.Key.Equals("NETFXKitsDir", StringComparison.OrdinalIgnoreCase)))
                properties.Add(new("NETFXKitsDir", netfxKitsDir));

            // Apply user-supplied --msbuild-property overrides last so they win over our defaults.
            foreach (var kv in _userProperties)
            {
                properties.RemoveAll(p => string.Equals(p.Key, kv.Key, StringComparison.OrdinalIgnoreCase));
                properties.Add(new(kv.Key, kv.Value));
            }

            // Don't set WindowsTargetPlatformVersion as a global property. It would
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

            // For .cmd/.bat MSBuild wrappers, use cmd.exe /c to ensure proper
            // batch script execution (init scripts, environment setup, etc.)
            // The default (auto) sniffs the file extension; users may override
            // via --msbuild-launcher cmd|direct|dotnet to force a specific mode.
            bool isCmdWrapper;
            switch (_launcher)
            {
                case MsBuildLauncher.Cmd:
                    isCmdWrapper = true;
                    break;
                case MsBuildLauncher.Direct:
                    isCmdWrapper = false;
                    break;
                case MsBuildLauncher.Dotnet:
                    // Force dotnet exec path; treat like a .dll (handled below).
                    isCmdWrapper = false;
                    break;
                default: // Auto
                    isCmdWrapper = _msbuildPath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
                                || _msbuildPath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase);
                    break;
            }

            if (_dotnetExePath != null || _launcher == MsBuildLauncher.Dotnet)
            {
                // MSBuild.dll: run via "dotnet exec MSBuild.dll ..."
                var dotnet = _dotnetExePath ?? FindDotnetExe()
                    ?? throw new FileNotFoundException("--msbuild-launcher dotnet was requested but 'dotnet' was not found on PATH.");
                startInfo.FileName = dotnet;
                startInfo.ArgumentList.Add("exec");
                startInfo.ArgumentList.Add(_msbuildPath);
            }
            else if (isCmdWrapper)
            {
                // For .cmd wrappers: build a single command string for cmd.exe /c
                // This ensures batch init scripts run correctly
                startInfo.FileName = "cmd.exe";
            }
            else
            {
                startInfo.FileName = _msbuildPath;
            }

            var msbuildArgs = new List<string>();
            msbuildArgs.Add(_projectPath);
            // Batch design-time targets. Even if some are unknown in older or custom .targets chains,
            // MSBuild skips missing targets with a warning rather than failing,
            // as long as they appear in a semicolon list after a known target.
            // ComputeReferenceCLInput is kept first so that referenced public include directories
            // are folded into ClCompile metadata before GetClCommandLines evaluates.
            msbuildArgs.Add("/t:ComputeReferenceCLInput;GetProjectDirectories;GetClCommandLines");

            // Use separate /p: per property to avoid issues with semicolons
            // or special characters in property values
            foreach (var prop in properties)
                msbuildArgs.Add($"/p:{prop.Key}={prop.Value}");

            msbuildArgs.Add($"/bl:{binlogPath}");
            msbuildArgs.Add("/nologo");
            msbuildArgs.Add("/v:quiet");

            if (isCmdWrapper)
            {
                // Build single command string for cmd.exe /c.
                var quotedArgs = msbuildArgs.Select(QuoteCmdArg);
                // Wrap the whole command in outer quotes: cmd /c "..." with multiple quoted
                // args inside. Per `cmd /?`, when more than two quotes are present cmd strips
                // the first and last quote and passes the rest through verbatim, so the outer
                // pair protects the inner quoting.
                startInfo.Arguments = $"/c \"\"{_msbuildPath}\" {string.Join(" ", quotedArgs)}\"";

                // When using a .cmd/.bat wrapper that launches .NET Framework MSBuild,
                // clear .NET SDK environment variables inherited from the dotnet host process.
                // These vars (MSBuildExtensionsPath, MSBUILD_EXE_PATH, etc.) cause .NET Framework
                // MSBuild to resolve targets from the .NET SDK instead of its own directory,
                // breaking target resolution (e.g., ResolveReferences not found).
                var dotnetSdkVarsToClean = new[]
                {
                    "MSBuildExtensionsPath",
                    "MSBuildSDKsPath",
                    "MSBUILD_EXE_PATH",
                    "MSBuildLoadMicrosoftTargetsReadOnly",
                    "MSBUILDFAILONDRIVEENUMERATINGWILDCARD",
                    "_MSBUILDTLENABLED",
                    "DOTNET_HOST_PATH",
                    "VCTargetsPath"
                };
                foreach (var varName in dotnetSdkVarsToClean)
                {
                    startInfo.Environment[varName] = "";
                }
            }
            else
            {
                foreach (var arg in msbuildArgs)
                    startInfo.ArgumentList.Add(arg);
            }

            // When the host process is a .NET SDK app (like this extractor),
            // clear .NET SDK environment variables that would cause .NET Framework MSBuild
            // to resolve targets from the wrong location.
            // This applies to both .cmd wrappers and direct MSBuild.exe launches.
            if (!isCmdWrapper)
            {
                var dotnetSdkVarsToClean = new[]
                {
                    "MSBuildExtensionsPath",
                    "MSBuildSDKsPath",
                    "MSBUILD_EXE_PATH",
                    "MSBuildLoadMicrosoftTargetsReadOnly",
                    "MSBUILDFAILONDRIVEENUMERATINGWILDCARD",
                    "_MSBUILDTLENABLED",
                    "DOTNET_HOST_PATH"
                };
                foreach (var varName in dotnetSdkVarsToClean)
                {
                    startInfo.Environment[varName] = "";
                }
            }

            // Set environment variables for non-.cmd cases
            if (!isCmdWrapper)
            {
                // Only set VCTargetsPath env var if not using MicrosoftBuildCppTasksCommonPath
                if (_vcTargetsPath != null && !properties.Any(p => p.Key == "MicrosoftBuildCppTasksCommonPath"))
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
            }

            // Apply user-supplied --msbuild-env overrides last so they win over our defaults
            // (including the .NET-SDK clear-list above and VCTargetsPath/VCToolsInstallDir).
            foreach (var kv in _userEnv)
            {
                startInfo.Environment[kv.Key] = kv.Value;
            }

            if (_enableLogger)
            {
                if (isCmdWrapper)
                {
                    Console.WriteLine($"Running: cmd.exe {startInfo.Arguments}");
                }
                else
                {
                    var cmdLine = _dotnetExePath != null
                        ? $"dotnet exec \"{_msbuildPath}\""
                        : $"\"{_msbuildPath}\"";
                    var args = string.Join(" ", startInfo.ArgumentList
                        .Skip(_dotnetExePath != null ? 2 : 0)
                        .Select(a => a.Contains(' ') ? $"\"{a}\"" : a));
                    Console.WriteLine($"Running: {cmdLine} {args}");
                }
            }

            using var process = System.Diagnostics.Process.Start(startInfo);
            if (process == null)
                throw new Exception("Failed to start MSBuild process");

            // Read both streams as tasks to avoid deadlock when both buffers fill
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            // Enforce timeout FIRST. ReadToEnd blocks forever if the process hangs
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

            // Prefer GetProjectDirectories target output over raw property reads,
            // it accounts for UseEnv=true and Makefile configuration resolution that
            // raw property values miss.
            string? includePath = null, externalIncludePath = null, vcToolsDirValue = null;
            var projDirsTarget = build.FindFirstDescendant<Target>(t => t.Name == "GetProjectDirectories");
            if (projDirsTarget != null)
            {
                var outputsFolder = projDirsTarget.FindFirstDescendant<Folder>(f => f.Name == "TargetOutputs");
                var directoriesItem = outputsFolder?.Children.OfType<Item>().FirstOrDefault();
                if (directoriesItem != null)
                {
                    foreach (var md in directoriesItem.Children.OfType<Metadata>())
                    {
                        switch (md.Name)
                        {
                            case "IncludePath": includePath = md.Value; break;
                            case "ExternalIncludePath": externalIncludePath = md.Value; break;
                        }
                    }
                }
            }

            // Fall back to Property nodes when GetProjectDirectories didn't run
            // (older targets, unknown-target warning).
            if (string.IsNullOrEmpty(externalIncludePath))
            {
                externalIncludePath = build.FindChildrenRecursive<Property>()
                    .LastOrDefault(p => p.Name == "ExternalIncludePath" && !string.IsNullOrEmpty(p.Value))?.Value;
            }
            if (string.IsNullOrEmpty(includePath))
            {
                includePath = build.FindChildrenRecursive<Property>()
                    .LastOrDefault(p => p.Name == "IncludePath" && !string.IsNullOrEmpty(p.Value))?.Value;
            }

            if (!string.IsNullOrEmpty(externalIncludePath))
            {
                _externalIncludePaths = ResolveIncludePathEntries(
                    externalIncludePath.Split(';', StringSplitOptions.RemoveEmptyEntries)).ToArray();
            }

            if (!string.IsNullOrEmpty(includePath))
            {
                var externalSet = new HashSet<string>(_externalIncludePaths, StringComparer.OrdinalIgnoreCase);
                _includePaths = ResolveIncludePathEntries(
                    includePath.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    .Where(p => !externalSet.Contains(p))
                    .ToArray();
            }

            // Add VS auxiliary include paths from VCToolsInstallDir
            vcToolsDirValue = build.FindChildrenRecursive<Property>()
                .LastOrDefault(p => p.Name == "VCToolsInstallDir" && !string.IsNullOrEmpty(p.Value))?.Value;
            var vcToolsDir = _vcToolsInstallDir ?? vcToolsDirValue?.TrimEnd('\\');
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

        private static bool DetectFxCompile(IReadOnlyList<string> files)
        {
            // The CLCommandLine task runs separately for @(ClCompile) and @(FxCompile),
            // and both outputs land in the same @(ClCommandLines) item group.
            // Detect HLSL by file extensions, since fxc.exe flags are not consumable by clangd,
            // so FxCompile entries must be dropped from standard output.
            if (files.Count == 0) return false;
            foreach (var f in files)
            {
                var ext = Path.GetExtension(f);
                if (!string.Equals(ext, ".hlsl", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(ext, ".fx", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(ext, ".hlsli", StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        private ClCommandLineEntry? ParseAddItem(AddItem addItem)
        {
            var commandLine = addItem.Name ?? "";
            var workingDir = "";
            var toolPath = "";
            var configOptions = "";
            var files = new List<string>();

            foreach (var metadata in addItem.Children.OfType<Metadata>())
            {
                switch (metadata.Name)
                {
                    case "WorkingDirectory": workingDir = metadata.Value; break;
                    case "ToolPath": toolPath = metadata.Value; break;
                    case "ConfigurationOptions": configOptions = metadata.Value; break;
                    case "Files":
                        files.AddRange(metadata.Value.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(f => f.Trim()));
                        break;
                }
            }

            if (files.Count == 0) return null;

            return new ClCommandLineEntry(
                CommandLine: commandLine,
                WorkingDirectory: workingDir,
                ToolPath: toolPath,
                Files: files.ToArray(),
                IsConfigurationDefault: string.Equals(configOptions, "true", StringComparison.OrdinalIgnoreCase),
                IsFxCompile: DetectFxCompile(files));
        }

        private ClCommandLineEntry? ParseItem(Item item)
        {
            var commandLine = item.Name ?? "";
            var workingDir = "";
            var toolPath = "";
            var configOptions = "";
            var files = new List<string>();

            foreach (var metadata in item.Children.OfType<Metadata>())
            {
                switch (metadata.Name)
                {
                    case "WorkingDirectory": workingDir = metadata.Value; break;
                    case "ToolPath": toolPath = metadata.Value; break;
                    case "ConfigurationOptions": configOptions = metadata.Value; break;
                    case "Files":
                        files.AddRange(metadata.Value.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(f => f.Trim()));
                        break;
                }
            }

            if (files.Count == 0) return null;

            return new ClCommandLineEntry(
                CommandLine: commandLine,
                WorkingDirectory: workingDir,
                ToolPath: toolPath,
                Files: files.ToArray(),
                IsConfigurationDefault: string.Equals(configOptions, "true", StringComparison.OrdinalIgnoreCase),
                IsFxCompile: DetectFxCompile(files));
        }

        private List<CompileCommand> ToCompileCommands(List<ClCommandLineEntry> entries)
        {
            var commands = new List<CompileCommand>();
            var projectName = Path.GetFileNameWithoutExtension(_projectPath);
            string[]? defaultArgs = null;

            foreach (var entry in entries)
            {
                // Drop FxCompile (HLSL) entries. fxc.exe flags are not
                // cl.exe-compatible and clangd cannot consume them.
                if (entry.IsFxCompile)
                    continue;

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

                // Decide where to place system include paths (from IncludePath / ExternalIncludePath MSBuild properties)
                // relative to the project's /I flags from AdditionalIncludeDirectories.
                //
                // The default (Auto) classifies each path:
                //   - paths inside the project tree are prepended,
                //     preserving the historical behavior for projects that place source dirs in IncludePath,
                //   - paths outside the project tree are appended,
                //     matching cl.exe semantics where INCLUDE env-var paths are searched after explicit /I,
                //     which avoids header shadowing when a system header has the same name as a generated header.
                var allSystemIncludes = _externalIncludePaths.Concat(_includePaths).ToArray();
                var (prependIncludes, appendIncludes) = ClassifyIncludePaths(allSystemIncludes);

                foreach (var p in prependIncludes)
                    baseArgs.Add($"/I{p}");

                baseArgs.AddRange(CommandLineTokenizer.TokenizeWithResponseFiles(entry.CommandLine, entry.WorkingDirectory));

                foreach (var p in appendIncludes)
                    baseArgs.Add($"/I{p}");

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
                    bool isTemporaryCpp = string.Equals(Path.GetFileName(file), "__temporary.cpp", StringComparison.OrdinalIgnoreCase);

                    if (isTemporaryCpp)
                    {
                        if (_emitDefaults)
                        {
                            var defaultsFile = Path.Combine(entry.WorkingDirectory, "__project_defaults.cpp");
                            var defaultsArgs = new List<string>(baseArgs) { defaultsFile };
                            var sanitized = InProcessExtractor.SanitizeFallbackPaths(defaultsArgs.ToArray(), _vcToolsInstallDir);
                            commands.Add(new CompileCommand(
                                File: defaultsFile,
                                Arguments: sanitized,
                                Directory: entry.WorkingDirectory,
                                ProjectPath: _projectPath,
                                ProjectName: projectName,
                                Configuration: _configuration,
                                Platform: _platform
                            ));
                        }
                        if (_mergeDefaults)
                            defaultArgs = baseArgs.ToArray();
                        continue;
                    }

                    var fileArgs = new List<string>(baseArgs) { file };

                    if (_mergeDefaults && defaultArgs != null)
                        InProcessExtractor.MergeDefaultFlags(fileArgs, defaultArgs);

                    var commandLine = InProcessExtractor.SanitizeFallbackPaths(fileArgs.ToArray(), _vcToolsInstallDir);

                    commands.Add(new CompileCommand(
                        File: file,
                        Arguments: commandLine,
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

        private (string[] prepend, string[] append) ClassifyIncludePaths(string[] all)
        {
            switch (_includePathOrder)
            {
                case IncludePathOrder.Prepend:
                    return (all, Array.Empty<string>());
                case IncludePathOrder.Append:
                    return (Array.Empty<string>(), all);
                default: // Auto: per-path classification by project-tree containment
                {
                    var projectDir = Path.GetDirectoryName(_projectPath);
                    if (string.IsNullOrEmpty(projectDir))
                        return (Array.Empty<string>(), all);

                    var projectDirNorm = NormalizePath(projectDir);
                    var prepend = new List<string>();
                    var append = new List<string>();
                    foreach (var p in all)
                    {
                        var n = NormalizePath(p);
                        if (n.StartsWith(projectDirNorm, StringComparison.OrdinalIgnoreCase))
                            prepend.Add(p);
                        else
                            append.Add(p);
                    }
                    return (prepend.ToArray(), append.ToArray());
                }
            }
        }

        private static string NormalizePath(string p)
        {
            try { return Path.GetFullPath(p).TrimEnd('\\') + "\\"; }
            catch { return p.TrimEnd('\\') + "\\"; }
        }

        private string ResolveToolPath(string toolPath)
        {
            // User-specified --cl-path takes highest priority
            if (!string.IsNullOrEmpty(_clPath) && File.Exists(_clPath))
                return _clPath;

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

            // Classify system include paths via the same policy as the main path.
            var allSystemIncludes2 = _externalIncludePaths.Concat(_includePaths).ToArray();
            var (prepend2, append2) = ClassifyIncludePaths(allSystemIncludes2);

            foreach (var p in prepend2)
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

            foreach (var p in append2)
                baseArgs.Add($"/I{p}");

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
                var commandLine = InProcessExtractor.SanitizeFallbackPaths(fileArgs.ToArray(), _vcToolsInstallDir);

                commands.Add(new CompileCommand(
                    File: fullPath,
                    Arguments: commandLine,
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
            string? vcToolsInstallDir = null, string? vcTargetsPath = null,
            IReadOnlyDictionary<string, string>? msbuildProperties = null,
            IReadOnlyDictionary<string, string>? msbuildEnv = null,
            MsBuildLauncher launcher = MsBuildLauncher.Auto,
            IncludePathOrder includePathOrder = IncludePathOrder.Auto,
            bool emitDefaults = false,
            bool mergeDefaults = false)
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
                        solutionDir, vcToolsInstallDir, vcTargetsPath,
                        msbuildProperties: msbuildProperties,
                        msbuildEnv: msbuildEnv,
                        launcher: launcher,
                        includePathOrder: includePathOrder,
                        emitDefaults: emitDefaults,
                        mergeDefaults: mergeDefaults);
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
