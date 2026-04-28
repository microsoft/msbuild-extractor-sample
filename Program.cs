using Microsoft.Build.Locator;
using System.Text.Json;

namespace MSBuild.CompileCommands.Extractor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var options = CommandLineOptions.Parse(args);

            // Handle --list-instances: print all VS installations and exit
            if (options.ListInstances)
            {
                var instances = VsWhereHelper.ListInstances();
                if (instances.Count == 0)
                {
                    Console.WriteLine("No Visual Studio installations found.");
                    return;
                }
                Console.WriteLine("Visual Studio Installations:");
                Console.WriteLine();
                foreach (var inst in instances)
                {
                    Console.WriteLine($"  ID:       {inst.InstanceId}");
                    Console.WriteLine($"  Name:     {inst.DisplayName}");
                    Console.WriteLine($"  Path:     {inst.InstallationPath}");
                    Console.WriteLine($"  Version:  {inst.Version}");
                    Console.WriteLine($"  VC Tools: {(inst.HasVcTools ? "Yes" : "No")}");
                    if (inst.VCTargetsPath != null)
                        Console.WriteLine($"  VCTargets: {inst.VCTargetsPath}");
                    if (inst.VCToolsInstallDir != null)
                        Console.WriteLine($"  VCTools:   {inst.VCToolsInstallDir}");
                    Console.WriteLine();
                }
                return;
            }

            // Handle --vs-instance: select a specific VS installation
            if (options.VsInstance != null)
            {
                var instances = VsWhereHelper.ListInstances(options.EnableLogger);
                var selected = instances.FirstOrDefault(i =>
                    i.InstanceId.Equals(options.VsInstance, StringComparison.OrdinalIgnoreCase) ||
                    i.InstanceId.StartsWith(options.VsInstance, StringComparison.OrdinalIgnoreCase));
                if (selected == null)
                {
                    Console.Error.WriteLine($"Error: VS instance '{options.VsInstance}' not found. Use --list-instances.");
                    Environment.Exit(1);
                }
                options.VsPath = selected.InstallationPath;
                if (options.VcTargetsPath == null) options.VcTargetsPath = selected.VCTargetsPath;
                if (options.VcToolsInstallDir == null) options.VcToolsInstallDir = selected.VCToolsInstallDir;
            }

            // Handle --use-dev-env: read environment variables from Developer Command Prompt
            if (options.UseDevEnv)
            {
                var devEnv = new DevEnvReader();
                if (options.EnableLogger)
                    devEnv.PrintDiagnostics();

                if (!devEnv.IsDevEnvAvailable)
                {
                    Console.Error.WriteLine("Error: --use-dev-env specified but no Developer Command Prompt environment detected.");
                    Console.Error.WriteLine("Run this tool from a Developer Command Prompt or Developer PowerShell.");
                    Environment.Exit(1);
                }

                if (options.VcTargetsPath == null && devEnv.VCTargetsPath != null)
                    options.VcTargetsPath = devEnv.VCTargetsPath;
                if (options.VcToolsInstallDir == null && devEnv.VCToolsInstallDir != null)
                    options.VcToolsInstallDir = devEnv.VCToolsInstallDir;

                devEnv.ApplyToEnvironment();
            }

            // Auto-detect via vswhere if VCTargetsPath or VCToolsInstallDir are still missing
            if (options.VcTargetsPath == null || options.VcToolsInstallDir == null)
            {
                var vsResult = VsWhereHelper.DetectVisualStudio(options.EnableLogger);
                if (vsResult != null)
                {
                    if (options.VcTargetsPath == null && vsResult.VCTargetsPath != null)
                        options.VcTargetsPath = vsResult.VCTargetsPath;
                    if (options.VcToolsInstallDir == null && vsResult.VCToolsInstallDir != null)
                        options.VcToolsInstallDir = vsResult.VCToolsInstallDir;
                    if (options.VsPath == null && vsResult.InstallationPath != null)
                        options.VsPath = vsResult.InstallationPath;
                }
            }

            // Register MSBuild early, needed for SolutionFile.Parse() in validation
            // and for in-process extraction. Out-of-process mode also needs it for
            // solution parsing even though extraction uses a separate MSBuild process.
            bool isOutOfProcess = options.MsBuildPath != null;

            if (options.EnableLogger)
                Console.WriteLine(isOutOfProcess
                    ? $"Mode: Out-of-process (MSBuild: {options.MsBuildPath})"
                    : "Mode: In-process");

            RegisterMSBuild(options.VsPath, options.EnableLogger);
            SetVcTargetsPath(options.VcTargetsPath);

            // Validate configuration/platform
            if (!options.AllConfigurations)
                ValidateConfigurations(options);

            bool isMultiInput = options.Solutions.Length + options.Projects.Length > 1;

            // Handle --all-configurations
            if (options.AllConfigurations)
            {
                RunAllConfigurations(options);
                return;
            }

            // Run extraction (always in-memory so we can prepend the sentinel entry)
            List<CompileCommand>? commands = null;
            if (isMultiInput)
            {
                commands = ExtractMultipleInputs(options);
            }
            else if (isOutOfProcess)
            {
                commands = ExtractOutOfProcess(options);
            }
            else
            {
                commands = ExtractInProcess(options);
            }

            // Write output when we collected commands in-memory
            if (commands != null)
            {
                if (options.Deduplicate)
                {
                    int beforeCount = commands.Count;
                    commands = CompileCommandDeduplicator.Deduplicate(
                        commands, options.PreferConfiguration, options.PreferPlatform);
                    Console.WriteLine($"Deduplicated: {beforeCount} entries → {commands.Count} unique files");
                }
                var outputPath = options.Output ?? GetDefaultOutputPath(options);
                WriteJson(commands, outputPath, options: options);
                Console.WriteLine($"Wrote {commands.Count} entries to {outputPath}");
                if (options.EmitCCppProperties)
                {
                    var baseDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".";
                    VsCodeSettings.GenerateCCppProperties(outputPath, options.Platform, baseDir);
                }
                if (options.Validate)
                    ClExeValidator.Run(commands);
            }
        }

        static void ValidateConfigurations(CommandLineOptions options)
        {
            foreach (var sln in options.Solutions)
            {
                var projects = ProjectDiscovery.GetVcProjectsFromSolution(sln);
                var errors = new List<string>();

                foreach (var project in projects)
                {
                    if (!project.HasConfigurationPlatform(options.Configuration, options.Platform))
                    {
                        var available = string.Join(", ", project.ConfigurationPlatforms.Select(cp => cp.ToString()));
                        errors.Add($"Configuration '{options.Configuration}|{options.Platform}' not found in {project.Name}. Available: {available}");
                    }
                }

                if (errors.Count > 0)
                {
                    foreach (var error in errors)
                    {
                        Console.Error.WriteLine(options.Strict ? $"Error: {error}" : $"Warning: {error}");
                    }
                    if (options.Strict)
                    {
                        Console.Error.WriteLine($"Aborting: {errors.Count} project(s) do not support the requested configuration.");
                        Environment.Exit(1);
                    }
                }
            }

            foreach (var proj in options.Projects)
            {
                var validationError = ProjectDiscovery.ValidateConfigurationPlatform(
                    proj, options.Configuration, options.Platform);
                if (validationError != null)
                {
                    if (options.Strict)
                    {
                        Console.Error.WriteLine($"Error: {validationError}");
                        Environment.Exit(1);
                    }
                    else
                    {
                        Console.Error.WriteLine($"Warning: {validationError}");
                    }
                }
            }
        }

        static void RunAllConfigurations(CommandLineOptions options)
        {
            var configPlatforms = DiscoverAllConfigurations(options);
            if (configPlatforms.Count == 0)
            {
                Console.Error.WriteLine("Error: No configuration/platform pairs found.");
                Environment.Exit(1);
            }

            Console.WriteLine($"Extracting for {configPlatforms.Count} configuration(s): {string.Join(", ", configPlatforms)}");

            bool isOutOfProcess = options.MsBuildPath != null;
            var allCommands = new List<CompileCommand>();
            var savedConfig = options.Configuration;
            var savedPlatform = options.Platform;

            foreach (var cp in configPlatforms)
            {
                Console.WriteLine($"  Extracting: {cp}...");
                options.Configuration = cp.Configuration;
                options.Platform = cp.Platform;

                try
                {
                    bool isMultiInput = options.Solutions.Length + options.Projects.Length > 1;
                    List<CompileCommand> commands;
                    if (isMultiInput)
                        commands = ExtractMultipleInputs(options);
                    else if (isOutOfProcess)
                        commands = ExtractOutOfProcess(options);
                    else
                        commands = ExtractInProcess(options);

                    if (options.Merge)
                    {
                        allCommands.AddRange(commands);
                    }
                    else
                    {
                        var outputDir = GetOutputDirectory(options);
                        var outputPath = Path.Combine(outputDir, $"compile_commands_{cp.Configuration}_{cp.Platform}.json");
                        WriteJson(commands, outputPath, options: options);
                        Console.WriteLine($"    Wrote {commands.Count} entries to {outputPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"    Warning: Failed for {cp}: {ex.Message}");
                }
            }

            options.Configuration = savedConfig;
            options.Platform = savedPlatform;

            if (options.Merge)
            {
                if (options.Deduplicate)
                {
                    int beforeCount = allCommands.Count;
                    allCommands = CompileCommandDeduplicator.Deduplicate(
                        allCommands, options.PreferConfiguration, options.PreferPlatform);
                    Console.WriteLine($"Deduplicated: {beforeCount} entries → {allCommands.Count} unique files");
                }
                var outputPath = options.Output ?? Path.Combine(GetOutputDirectory(options), "compile_commands.json");
                WriteJson(allCommands, outputPath, includeMetadata: !options.Deduplicate, options: options);
                Console.WriteLine($"Wrote {allCommands.Count} entries{(options.Deduplicate ? " (deduplicated)" : " (merged)")} to {outputPath}");
                if (options.EmitCCppProperties)
                {
                    var baseDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".";
                    VsCodeSettings.GenerateCCppProperties(outputPath, options.Platform, baseDir);
                }
            }
        }

        static List<ConfigurationPlatform> DiscoverAllConfigurations(CommandLineOptions options)
        {
            var all = new List<ConfigurationPlatform>();

            foreach (var sln in options.Solutions)
            {
                all.AddRange(ProjectDiscovery.GetVcProjectsFromSolution(sln)
                    .SelectMany(p => p.ConfigurationPlatforms));
            }

            foreach (var proj in options.Projects)
            {
                all.AddRange(ProjectDiscovery.GetProjectConfigurations(proj));
            }

            return all.Distinct().ToList();
        }

        static string GetOutputDirectory(CommandLineOptions options)
        {
            if (options.Output != null)
                return Path.GetDirectoryName(Path.GetFullPath(options.Output)) ?? ".";
            var inputPath = options.Solutions.FirstOrDefault() ?? options.Projects.First();
            return Path.GetDirectoryName(Path.GetFullPath(inputPath)) ?? ".";
        }

        static List<CompileCommand> ExtractInProcess(CommandLineOptions options)
        {
            if (options.Solutions.Length > 0 && options.Solution != null)
            {
                return InProcessExtractor.ExtractCompileCommandsFromSolution(
                    options.Solution, options.Configuration, options.Platform,
                    options.EnableLogger, options.VcToolsInstallDir,
                    options.EmitDefaults, options.MergeDefaults);
            }
            else
            {
                var extractor = new InProcessExtractor(
                    options.Projects[0], options.Configuration, options.Platform,
                    options.EnableLogger, options.SolutionDir, options.VcToolsInstallDir,
                    options.EmitDefaults, options.MergeDefaults);
                return extractor.ExtractCompileCommands();
            }
        }

        static List<CompileCommand> ExtractOutOfProcess(CommandLineOptions options)
        {
            if (options.Solutions.Length > 0 && options.Solution != null)
            {
                return OutOfProcessExtractor.ExtractCompileCommandsFromSolution(
                    options.MsBuildPath!, options.Solution, options.Configuration, options.Platform,
                    options.EnableLogger, options.VcToolsInstallDir, options.VcTargetsPath,
                    emitDefaults: options.EmitDefaults, mergeDefaults: options.MergeDefaults);
            }
            else
            {
                var extractor = new OutOfProcessExtractor(
                    options.MsBuildPath!, options.Projects[0], options.Configuration, options.Platform,
                    options.EnableLogger, options.SolutionDir, options.VcToolsInstallDir, options.VcTargetsPath,
                    options.ClPath,
                    msbuildProperties: options.MsBuildProperties,
                    msbuildEnv: options.MsBuildEnv,
                    launcher: ParseLauncher(options.MsBuildLauncher),
                    includePathOrder: ParseIncludePathOrder(options.IncludePathOrder),
                    emitDefaults: options.EmitDefaults,
                    mergeDefaults: options.MergeDefaults);
                return extractor.ExtractCompileCommands();
            }
        }

        static List<CompileCommand> ExtractMultipleInputs(CommandLineOptions options)
        {
            bool isOutOfProcess = options.MsBuildPath != null;
            var allCommands = new List<CompileCommand>();

            foreach (var sln in options.Solutions)
            {
                Console.WriteLine($"Extracting from solution: {Path.GetFileName(sln)}...");
                try
                {
                    List<CompileCommand> commands;
                    if (isOutOfProcess)
                        commands = OutOfProcessExtractor.ExtractCompileCommandsFromSolution(
                            options.MsBuildPath!, sln, options.Configuration, options.Platform,
                            options.EnableLogger, options.VcToolsInstallDir, options.VcTargetsPath,
                            msbuildProperties: options.MsBuildProperties,
                            msbuildEnv: options.MsBuildEnv,
                            launcher: ParseLauncher(options.MsBuildLauncher),
                            includePathOrder: ParseIncludePathOrder(options.IncludePathOrder),
                            emitDefaults: options.EmitDefaults,
                            mergeDefaults: options.MergeDefaults);
                    else
                        commands = InProcessExtractor.ExtractCompileCommandsFromSolution(
                            sln, options.Configuration, options.Platform,
                            options.EnableLogger, options.VcToolsInstallDir,
                            options.EmitDefaults, options.MergeDefaults);
                    Console.WriteLine($"  Got {commands.Count} entries");
                    allCommands.AddRange(commands);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"  Warning: Failed for {sln}: {ex.Message}");
                }
            }

            foreach (var proj in options.Projects)
            {
                Console.WriteLine($"Extracting from project: {Path.GetFileName(proj)}...");
                try
                {
                    List<CompileCommand> commands;
                    if (isOutOfProcess)
                    {
                        var extractor = new OutOfProcessExtractor(
                            options.MsBuildPath!, proj, options.Configuration, options.Platform,
                            options.EnableLogger, options.SolutionDir, options.VcToolsInstallDir, options.VcTargetsPath,
                            options.ClPath,
                            msbuildProperties: options.MsBuildProperties,
                            msbuildEnv: options.MsBuildEnv,
                            launcher: ParseLauncher(options.MsBuildLauncher),
                            includePathOrder: ParseIncludePathOrder(options.IncludePathOrder),
                            emitDefaults: options.EmitDefaults,
                            mergeDefaults: options.MergeDefaults);
                        commands = extractor.ExtractCompileCommands();
                    }
                    else
                    {
                        var extractor = new InProcessExtractor(
                            proj, options.Configuration, options.Platform,
                            options.EnableLogger, options.SolutionDir, options.VcToolsInstallDir,
                            options.EmitDefaults, options.MergeDefaults);
                        commands = extractor.ExtractCompileCommands();
                    }
                    Console.WriteLine($"  Got {commands.Count} entries");
                    allCommands.AddRange(commands);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"  Warning: Failed for {proj}: {ex.Message}");
                }
            }

            return allCommands;
        }

        static void WriteJson(List<CompileCommand> commands, string outputPath, bool includeMetadata = false,
            CommandLineOptions? options = null)
        {
            // Rich format: hierarchical compile_database.json
            if (options?.Format == "rich")
            {
                var root = RichDatabase.Build(commands, options);
                File.WriteAllText(outputPath, RichDatabase.Serialize(root));
                return;
            }

            // Prepend a sentinel entry so consumers can identify this file as generated
            // by msbuild-extractor-sample. Tools like clangd silently skip entries whose
            // file does not exist on disk, so this is invisible to all standard consumers.
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".";
            var sentinel = new
            {
                file = ".msbuild-extractor-sample",
                directory = outputDir,
                command = "generated-by:msbuild-extractor-sample/1.0.0"
            };

            string json;
            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            if (includeMetadata)
            {
                var entries = commands.Select(c => (object)new
                {
                    file = c.File,
                    arguments = c.Arguments,
                    directory = c.Directory,
                    projectPath = c.ProjectPath,
                    projectName = c.ProjectName,
                    configuration = c.Configuration,
                    platform = c.Platform
                });
                var all = new object[] { sentinel }.Concat(entries).ToArray();
                json = JsonSerializer.Serialize(all, jsonOptions);
            }
            else
            {
                var entries = commands.Select(c => (object)new
                {
                    file = c.File,
                    arguments = c.Arguments,
                    directory = c.Directory
                });
                var all = new object[] { sentinel }.Concat(entries).ToArray();
                json = JsonSerializer.Serialize(all, jsonOptions);
            }
            File.WriteAllText(outputPath, json);
        }

        static void SetVcTargetsPath(string? vcTargetsPath)
        {
            if (vcTargetsPath != null)
            {
                var path = vcTargetsPath.EndsWith('\\') ? vcTargetsPath : vcTargetsPath + "\\";
                Environment.SetEnvironmentVariable("VCTargetsPath", path);
            }
        }

        static string GetDefaultOutputPath(CommandLineOptions options)
        {
            var inputPath = options.Solutions.FirstOrDefault() ?? options.Projects.First();
            var dir = Path.GetDirectoryName(Path.GetFullPath(inputPath))!;
            var filename = options.Format == "rich" ? "compile_database.json" : "compile_commands.json";
            return Path.Combine(dir, filename);
        }

        static void RegisterMSBuild(string? vsPath, bool enableLogger = false)
        {
            if (vsPath != null)
            {
                var possiblePaths = new[]
                {
                    Path.Combine(vsPath, "MSBuild", "Current", "Bin", "amd64"),
                    Path.Combine(vsPath, "MSBuild", "Current", "Bin"),
                    Path.Combine(vsPath, "MSBuild", "Current", "Bin", "x86"),
                };

                foreach (var msbuildPath in possiblePaths)
                {
                    if (Directory.Exists(msbuildPath) && File.Exists(Path.Combine(msbuildPath, "MSBuild.dll")))
                    {
                        if (enableLogger)
                            Console.WriteLine($"MSBuild: Using --vs-path: {msbuildPath}");
                        MSBuildLocator.RegisterMSBuildPath(msbuildPath);
                        return;
                    }
                }

                Console.Error.WriteLine($"Warning: Could not find MSBuild in {vsPath}, falling back to discovery");
            }

            var instances = MSBuildLocator.QueryVisualStudioInstances()
                .Where(i => i.DiscoveryType == DiscoveryType.VisualStudioSetup)
                .OrderByDescending(i => i.Version)
                .ToList();

            if (enableLogger)
            {
                Console.WriteLine("MSBuild Discovery:");
                foreach (var inst in MSBuildLocator.QueryVisualStudioInstances())
                {
                    var marker = instances.Count > 0 && inst == instances[0] ? " <-- selected" : "";
                    Console.WriteLine($"  [{inst.DiscoveryType}] {inst.Name} v{inst.Version} @ {inst.MSBuildPath}{marker}");
                }
                Console.WriteLine();
            }

            if (instances.Count > 0)
                MSBuildLocator.RegisterInstance(instances[0]);
            else
                MSBuildLocator.RegisterDefaults();
        }

        private static MsBuildLauncher ParseLauncher(string s) => s.ToLowerInvariant() switch
        {
            "cmd" => MsBuildLauncher.Cmd,
            "direct" => MsBuildLauncher.Direct,
            "dotnet" => MsBuildLauncher.Dotnet,
            _ => MsBuildLauncher.Auto
        };

        private static IncludePathOrder ParseIncludePathOrder(string s) => s.ToLowerInvariant() switch
        {
            "prepend" => IncludePathOrder.Prepend,
            "append" => IncludePathOrder.Append,
            _ => IncludePathOrder.Auto
        };
    }
}