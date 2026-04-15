using Microsoft.Build.Execution;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MSBuild.CompileCommands.Extractor
{
    public class InProcessExtractor : ICompileCommandsExtractor
    {
        private readonly string _projectPath;
        private readonly string _projectName;
        private readonly bool _enableLogger;
        private readonly string _configuration;
        private readonly string _platform;
        private readonly Dictionary<string, string> _globalProperties;
        private readonly string? _vcToolsInstallDir;
        private string[] _externalIncludePaths = [];
        private string[] _includePaths = [];
        private string? _fixupPropsPath;

        // Resolve unevaluated MSBuild expressions like $([MSBuild]::NormalizePath('base', 'rel'))
        // that appear in IncludePath when GetPropertyValue doesn't fully evaluate intrinsics.
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

        // Derive VS auxiliary include paths from VCToolsInstallDir.
        // VCToolsInstallDir is like ...\VC\Tools\MSVC\14.50.35717
        // VS auxiliary paths are ...\VC\Auxiliary\VS\include and ...\VC\Auxiliary\VS\UnitTest\include
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

        public InProcessExtractor(
            string projectPath,
            string configuration = "Debug",
            string platform = "x64",
            bool enableLogger = false,
            string? solutionDir = null,
            string? vcToolsInstallDir = null)
        {
            _projectPath = Path.GetFullPath(projectPath);

            if (!File.Exists(_projectPath))
                throw new FileNotFoundException($"Project file not found: {_projectPath}");

            _projectName = Path.GetFileNameWithoutExtension(_projectPath);
            _enableLogger = enableLogger;
            _configuration = configuration;
            _platform = platform;
            _vcToolsInstallDir = vcToolsInstallDir?.TrimEnd('\\');
            if (_vcToolsInstallDir != null && !_vcToolsInstallDir.StartsWith(@"\\"))
                _vcToolsInstallDir = _vcToolsInstallDir.Replace(@"\\", @"\");

            _globalProperties = new Dictionary<string, string>
            {
                { "Configuration", configuration },
                { "Platform", platform },
                { "DesignTimeBuild", "true" },
                { "BuildingInsideVisualStudio", "true" },
                { "BuildProjectReferences", "true" },
                // VC++ projects use OutDir, not OutputPath. When the .NET SDK's
                // _CheckForInvalidConfigurationAndPlatform target runs, it fails
                // if OutputPath is empty. SkipInvalidConfigurations bypasses this.
                { "SkipInvalidConfigurations", "true" }
            };

            // Create a temporary .props file that defines default ClCompile metadata.
            // FixupCLCompileOptions in Microsoft.CppCommon.Targets uses unqualified
            // %(PrecompiledHeaderFile) for batching, causing MSB4096 when some items
            // have the metadata and others don't. ForceImportBeforeCppTargets injects
            // the defaults at the XML import level before targets execute.
            _fixupPropsPath = CreateFixupPropsFile();
            if (_fixupPropsPath != null)
            {
                _globalProperties["ForceImportBeforeCppTargets"] = _fixupPropsPath;
            }

            if (solutionDir is not null)
            {
                _globalProperties["SolutionDir"] = solutionDir.EndsWith('\\')
                    ? solutionDir : solutionDir + "\\";
            }

            if (vcToolsInstallDir is not null)
            {
                var normalized = vcToolsInstallDir.StartsWith(@"\\")
                    ? vcToolsInstallDir
                    : vcToolsInstallDir.Replace(@"\\", @"\");
                _globalProperties["VCToolsInstallDir"] = normalized.EndsWith('\\')
                    ? normalized : normalized + "\\";
            }

            // WindowsTargetPlatformVersion is resolved after project evaluation
            // so that projects with explicit full versions aren't overridden.
        }

        public List<CompileCommand> ExtractCompileCommands()
        {
            try
            {
                List<ClCommandLineEntry> entries;
                try
                {
                    entries = GetClCommandLines();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Info: Design-time build failed for {_projectPath}: {ex.Message}");
                    entries = new List<ClCommandLineEntry>();
                }

                var commands = entries.Count > 0 ? ToCompileCommands(entries) : new List<CompileCommand>();
                if (commands.Count > 0)
                    return commands;

                // Fallback for GN-generated projects (Chromium, Crashpad, WebRTC):
                // These have IntelliSense settings in ItemDefinitionGroup/ClCompile
                // but source files as CustomBuild items that delegate to ninja.
                var fallback = TryExtractFromItemDefinitionGroup();
                if (fallback.Count > 0)
                {
                    Console.Error.WriteLine($"Info: Used ItemDefinitionGroup fallback for {_projectPath} ({fallback.Count} entries)");
                    return fallback;
                }

                Console.Error.WriteLine($"Warning: No compile commands found in {_projectPath}");
                return new List<CompileCommand>();
            }
            finally
            {
                CleanupFixupPropsFile();
            }
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
            var projectInstance = new ProjectInstance(_projectPath, _globalProperties, null);

            // Resolve ambiguous WindowsTargetPlatformVersion after evaluation
            // so projects with explicit full versions (e.g. 10.0.19041.0) aren't overridden.
            var wtp = projectInstance.GetPropertyValue("WindowsTargetPlatformVersion");
            if (string.IsNullOrEmpty(wtp) || wtp == "10" || wtp == "10.0")
            {
                var resolved = VsWhereHelper.FindLatestWindowsSdkVersion();
                if (resolved != null)
                    projectInstance.SetProperty("WindowsTargetPlatformVersion", resolved);
            }

            // Fix MSB4096: FixupCLCompileOptions in Microsoft.CppCommon.Targets uses
            // unqualified %(PrecompiledHeaderFile) for batching. When some ClCompile items
            // have this metadata (from ItemDefinitionGroup) and others don't, MSBuild
            // throws MSB4096. Ensure all items have the metadata defined (even if empty).
            EnsureClCompileMetadataConsistency(projectInstance);

            var externalIncludes = projectInstance.GetPropertyValue("ExternalIncludePath");
            if (!string.IsNullOrEmpty(externalIncludes))
            {
                _externalIncludePaths = ResolveIncludePathEntries(
                    externalIncludes.Split(';', StringSplitOptions.RemoveEmptyEntries)).ToArray();
            }

            // IncludePath is a project-level property that feeds the compiler's
            // system include directories (like the INCLUDE env var). Projects such
            // as SDL add source directories here instead of AdditionalIncludeDirectories.
            // GetClCommandLines does not emit /I flags for these, so we inject them.
            var includePath = projectInstance.GetPropertyValue("IncludePath");
            if (!string.IsNullOrEmpty(includePath))
            {
                var externalSet = new HashSet<string>(
                    _externalIncludePaths, StringComparer.OrdinalIgnoreCase);
                _includePaths = ResolveIncludePathEntries(
                    includePath.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    .Where(p => !externalSet.Contains(p))
                    .ToArray();
            }

            // Add VS auxiliary include paths (VC\Auxiliary\VS\include, VC\Auxiliary\VS\UnitTest\include)
            // that come from $([MSBuild]::NormalizePath(...)) in IncludePath which can't be fully
            // evaluated by GetPropertyValue due to nested $(MSBuildThisFileDirectory) references.
            var vcToolsDir = _vcToolsInstallDir ?? projectInstance.GetPropertyValue("VCToolsInstallDir")?.TrimEnd('\\');
            var allIncludes = new HashSet<string>(_externalIncludePaths.Concat(_includePaths), StringComparer.OrdinalIgnoreCase);
            var auxPaths = GetVsAuxiliaryIncludePaths(vcToolsDir)
                .Where(p => !allIncludes.Contains(p))
                .ToArray();
            if (auxPaths.Length > 0)
                _includePaths = _includePaths.Concat(auxPaths).ToArray();

            var buildParams = new BuildParameters();
            if (_enableLogger)
            {
                buildParams.Loggers = new[] { new Microsoft.Build.Logging.ConsoleLogger(Microsoft.Build.Framework.LoggerVerbosity.Detailed) };
            }

            var buildRequestData = new BuildRequestData(projectInstance, new[] { "GetClCommandLines" });
            var buildResult = BuildManager.DefaultBuildManager.Build(buildParams, buildRequestData);

            if (buildResult.OverallResult != BuildResultCode.Success)
            {
                // GN-generated projects may fail because the configuration doesn't match.
                // Return empty list to trigger the ItemDefinitionGroup fallback.
                if (_enableLogger)
                    Console.Error.WriteLine($"Design-time build failed for {_projectPath} (will try fallback)");
                return new List<ClCommandLineEntry>();
            }

            var entries = new List<ClCommandLineEntry>();
            var targetResult = buildResult.ResultsByTarget["GetClCommandLines"];

            foreach (var item in targetResult.Items)
            {
                entries.Add(new ClCommandLineEntry(
                    CommandLine: item.ItemSpec,
                    WorkingDirectory: item.GetMetadata("WorkingDirectory"),
                    ToolPath: item.GetMetadata("ToolPath"),
                    Files: item.GetMetadata("Files")
                        .Split(';', StringSplitOptions.RemoveEmptyEntries)
                        .Select(f => f.Trim())
                        .ToArray()
                ));
            }

            return entries;
        }

        /// <summary>
        /// Ensures metadata consistency across ClCompile items to prevent MSB4096 errors.
        /// The VC++ targets use unqualified %(PrecompiledHeaderFile) for batching, which
        /// fails when some items define it and others don't. Projects like Windows Calculator
        /// set PrecompiledHeader=Use in ItemDefinitionGroup but omit PrecompiledHeaderFile,
        /// while the PCH source file has PrecompiledHeader=Create set per-item.
        /// </summary>
        private static void EnsureClCompileMetadataConsistency(ProjectInstance projectInstance)
        {
            var clCompileItems = projectInstance.GetItems("ClCompile").ToList();
            if (clCompileItems.Count == 0) return;

            // Metadata keys that FixupCLCompileOptions batches on without qualification
            var metadataKeys = new[] { "PrecompiledHeaderFile", "PrecompiledHeaderOutputFile" };

            foreach (var key in metadataKeys)
            {
                bool anyHasValue = clCompileItems.Any(i => !string.IsNullOrEmpty(i.GetMetadataValue(key)));
                if (!anyHasValue) continue;

                // Some items have it, some don't — set empty string on items missing it
                foreach (var item in clCompileItems)
                {
                    if (string.IsNullOrEmpty(item.GetMetadataValue(key)))
                    {
                        item.SetMetadata(key, "");
                    }
                }
            }
        }

        private static readonly HashSet<string> CppExtensions = new(StringComparer.OrdinalIgnoreCase)
            { ".cpp", ".cc", ".cxx", ".c", ".c++", ".cp" };

        /// <summary>
        /// Fallback extraction for GN-generated projects (Chromium, Crashpad, WebRTC).
        /// These projects store IntelliSense settings in ItemDefinitionGroup/ClCompile
        /// and list source files as CustomBuild items that delegate compilation to ninja.
        /// Parses the vcxproj XML directly to avoid MSBuild evaluation issues with
        /// non-standard configurations (e.g., "GN" instead of "Debug").
        /// </summary>
        private List<CompileCommand> TryExtractFromItemDefinitionGroup()
        {
            XDocument doc;
            try
            {
                doc = XDocument.Load(_projectPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: Failed to load project XML for fallback: {ex.Message}");
                return new List<CompileCommand>();
            }

            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var projectDir = Path.GetDirectoryName(_projectPath)!;

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
            // files present on disk nearby (e.g., wrapper executables, header-only
            // libs with associated .cc files).
            if (sourceFiles.Count == 0)
            {
                sourceFiles = FindSourceFilesOnDisk(projectDir, doc, ns);
            }

            if (sourceFiles.Count == 0)
                return new List<CompileCommand>();

            // Build the compiler arguments from ItemDefinitionGroup metadata
            var toolPath = ResolveToolPath(VsWhereHelper.ResolveClExePath(_vcToolsInstallDir, _platform)
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "CL.exe"));

            var baseArgs = new List<string> { toolPath };

            var target = GetClangTargetTriple();
            if (target != null)
                baseArgs.Add($"--target={target}");
            baseArgs.Add("-ferror-limit=0");

            // Add /I from ExternalIncludePath and IncludePath properties
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

            // Add AdditionalOptions (raw flags, may contain /D, /std:, warnings, etc.)
            if (!string.IsNullOrEmpty(additionalOptions))
            {
                var tokens = CommandLineTokenizer.TokenizeWithResponseFiles(
                    additionalOptions, projectDir);
                // Filter out %(AdditionalOptions) MSBuild inheritance marker
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

            // Build compile commands for each source file
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
                    ProjectName: _projectName,
                    Configuration: _configuration,
                    Platform: _platform
                ));
            }

            return commands;
        }

        /// <summary>
        /// Scans the disk for C/C++ source files when the project XML contains none.
        /// Tries: ClInclude/None item directories first, then the project directory itself.
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
                // Scan directories where headers were found for matching source files
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
                // No header hints — scan the project directory (non-recursive to avoid
                // pulling in huge trees for build-output directories)
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
        /// Creates a temporary .props file that defines default ItemDefinitionGroup metadata
        /// for ClCompile items, preventing MSB4096 errors from unqualified metadata batching
        /// in FixupCLCompileOptions. This works at the MSBuild XML import level.
        /// </summary>
        private static string? CreateFixupPropsFile()
        {
            try
            {
                var propsPath = Path.Combine(Path.GetTempPath(), $"msbuild_extractor_fixup_{Guid.NewGuid():N}.props");
                var propsContent = """
                    <Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
                      <ItemDefinitionGroup>
                        <ClCompile>
                          <PrecompiledHeaderFile Condition="'%(ClCompile.PrecompiledHeaderFile)' == ''">pch.h</PrecompiledHeaderFile>
                          <PrecompiledHeaderOutputFile Condition="'%(ClCompile.PrecompiledHeaderOutputFile)' == ''">$(IntDir)%(Filename).pch</PrecompiledHeaderOutputFile>
                        </ClCompile>
                      </ItemDefinitionGroup>
                    </Project>
                    """;
                File.WriteAllText(propsPath, propsContent);
                return propsPath;
            }
            catch
            {
                return null;
            }
        }

        private void CleanupFixupPropsFile()
        {
            if (_fixupPropsPath != null)
            {
                try { File.Delete(_fixupPropsPath); } catch { }
                _fixupPropsPath = null;
            }
        }

        private List<CompileCommand> ToCompileCommands(List<ClCommandLineEntry> entries)
        {
            var commands = new List<CompileCommand>();

            foreach (var entry in entries)
            {
                var toolPath = ResolveToolPath(entry.ToolPath);
                var baseArgs = new List<string> { toolPath };

                // Inject --target so clangd uses the correct architecture triple.
                // cl.exe path encodes target arch (e.g. Hostx64\x86) but clangd
                // defaults to the host triple (x86_64) with --driver-mode=cl.
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
                        ProjectName: _projectName,
                        Configuration: _configuration,
                        Platform: _platform
                    ));
                }
            }

            return commands;
        }

        /// <summary>
        /// Maps the MSBuild Platform to a clang target triple so that clangd
        /// compiles with the correct pointer size, predefined macros (_WIN64,
        /// _M_X64, etc.), and ABI. Without this, Win32 projects that define
        /// _USE_32BIT_TIME_T break because clangd defaults to x86_64.
        /// </summary>
        private string? GetClangTargetTriple()
        {
            return _platform.ToLowerInvariant() switch
            {
                "win32" or "x86" => "i686-pc-windows-msvc",
                "x64" or "x86_64" or "amd64" => "x86_64-pc-windows-msvc",
                "arm" => "thumbv7-pc-windows-msvc",
                "arm64" or "aarch64" => "aarch64-pc-windows-msvc",
                _ => null // Let clangd infer from cl.exe path
            };
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

        public static List<CompileCommand> ExtractCompileCommandsFromSolution(
            string solutionPath, string configuration = "Debug", string platform = "x64",
            bool enableLogger = false, string? vcToolsInstallDir = null)
        {
            var solutionDir = Path.GetDirectoryName(Path.GetFullPath(solutionPath))!;
            var projects = ProjectDiscovery.GetVcProjectsFromSolution(solutionPath);
            var allCommands = new List<CompileCommand>();

            foreach (var project in projects)
            {
                try
                {
                    var extractor = new InProcessExtractor(
                        project.Path, configuration, platform, enableLogger, solutionDir, vcToolsInstallDir);
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
            string solutionPath, string configuration = "Debug", string platform = "x64",
            bool enableLogger = false, string? outputPath = null, string? vcToolsInstallDir = null)
        {
            var commands = ExtractCompileCommandsFromSolution(solutionPath, configuration, platform, enableLogger, vcToolsInstallDir);
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
