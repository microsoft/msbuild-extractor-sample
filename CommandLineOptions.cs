using System.CommandLine;

namespace MSBuild.CompileCommands.Extractor
{
    public class CommandLineOptions
    {
        public string[] Projects { get; set; } = [];
        public string[] Solutions { get; set; } = [];

        public string? Project => Projects.Length == 1 ? Projects[0] : null;
        public string? Solution => Solutions.Length == 1 ? Solutions[0] : null;
        public string Configuration { get; set; } = "Debug";
        public string Platform { get; set; } = "x64";
        public string? VsPath { get; set; }
        public string? VcTargetsPath { get; set; }
        public string? ClPath { get; set; }
        public string? SolutionDir { get; set; }
        public string? VcToolsInstallDir { get; set; }
        public string? MsBuildPath { get; set; }
        public string? Output { get; set; }
        public bool EnableLogger { get; set; }
        public bool UseDevEnv { get; set; }
        public bool AllConfigurations { get; set; }
        public bool Merge { get; set; }
        public bool Strict { get; set; }
        public bool Validate { get; set; }
        public string Format { get; set; } = "standard";
        public bool Deduplicate { get; set; }
        public string PreferConfiguration { get; set; } = "Debug";
        public string PreferPlatform { get; set; } = "x64";
        public bool ListInstances { get; set; }
        public string? VsInstance { get; set; }
        public bool EmitCCppProperties { get; set; }
        public bool EmitDefaults { get; set; }
        public bool MergeDefaults { get; set; }
        public Dictionary<string, string> MsBuildProperties { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> MsBuildEnv { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public string MsBuildLauncher { get; set; } = "auto";
        public string IncludePathOrder { get; set; } = "auto";

        /// <summary>Path to the JSON config file that supplied defaults, or null if none was used.</summary>
        public string? ConfigPath { get; set; }

        public static CommandLineOptions Parse(string[] args)
        {
            CommandLineOptions? result = null;

            // Load an optional config file; its values become the defaults for the options built
            // below, so anything passed on the command line still wins. config stays null when no
            // config file is present.
            Config? config = new Config();
            if (!config.Load(args))
                config = null;

            var projectOption = new Option<string[]>("--project", "-p")
            {
                Description = "Path to .vcxproj file (can be specified multiple times)",
                AllowMultipleArgumentsPerToken = true,
                DefaultValueFactory = _ => config?.Projects ?? []
            };

            var solutionOption = new Option<string[]>("--solution", "-s")
            {
                Description = "Path to .sln or .slnx file (can be specified multiple times)",
                AllowMultipleArgumentsPerToken = true,
                DefaultValueFactory = _ => config?.Solutions ?? []
            };

            var configOption = new Option<string>("--configuration", "-c")
            {
                Description = "Build configuration",
                DefaultValueFactory = _ => config?.Configuration ?? "Debug"
            };

            var platformOption = new Option<string>("--platform", "-a")
            {
                Description = "Build platform",
                DefaultValueFactory = _ => config?.Platform ?? "x64"
            };

            var vsPathOption = new Option<string?>("--vs-path")
            {
                Description = "Path to Visual Studio installation",
                DefaultValueFactory = _ => config?.VsPath
            };

            var vcTargetsPathOption = new Option<string?>("--vc-targets-path")
            {
                Description = "Path to VC targets (e.g. ...\\MSBuild\\Microsoft\\VC\\v180)",
                DefaultValueFactory = _ => config?.VcTargetsPath
            };

            var clPathOption = new Option<string?>("--cl-path")
            {
                Description = "Path to cl.exe",
                DefaultValueFactory = _ => config?.ClPath
            };

            var solutionDirOption = new Option<string?>("--solution-dir")
            {
                Description = "Value for the SolutionDir MSBuild property (auto-derived when using --solution)",
                DefaultValueFactory = _ => config?.SolutionDir
            };

            var vcToolsInstallDirOption = new Option<string?>("--vc-tools-install-dir")
            {
                Description = "Value for the VCToolsInstallDir MSBuild property",
                DefaultValueFactory = _ => config?.VcToolsInstallDir
            };

            var outputOption = new Option<string?>("--output", "-o")
            {
                Description = "Output path for compile_commands.json",
                DefaultValueFactory = _ => config?.Output
            };

            var loggerOption = new Option<bool>("--logger")
            {
                Description = "Enable MSBuild console logger output",
                DefaultValueFactory = _ => config?.Logger ?? false
            };

            var useDevEnvOption = new Option<bool>("--use-dev-env")
            {
                Description = "Read environment variables from Developer Command Prompt (VCToolsInstallDir, VCTargetsPath, etc.)",
                DefaultValueFactory = _ => config?.UseDevEnv ?? false
            };

            var msbuildPathOption = new Option<string?>("--msbuild-path")
            {
                Description = "Path to msbuild.exe (enables out-of-process mode for custom MSBuild/toolchain locations)",
                DefaultValueFactory = _ => config?.MsBuildPath
            };

            var allConfigsOption = new Option<bool>("--all-configurations")
            {
                Description = "Extract for all configuration/platform combinations found in the project or solution",
                DefaultValueFactory = _ => config?.AllConfigurations ?? false
            };

            var mergeOption = new Option<bool>("--merge")
            {
                Description = "Merge all configurations into a single output file (use with --all-configurations)",
                DefaultValueFactory = _ => config?.Merge ?? false
            };

            var strictOption = new Option<bool>("--strict")
            {
                Description = "Treat configuration/platform validation warnings as errors",
                DefaultValueFactory = _ => config?.Strict ?? false
            };

            var validateOption = new Option<bool>("--validate")
            {
                Description = "After extraction, verify each compile command by running cl.exe /c",
                DefaultValueFactory = _ => config?.Validate ?? false
            };

            var formatOption = new Option<string>("--format", "-f")
            {
                Description = "Output format: 'standard' (compile_commands.json) or 'rich' (hierarchical compile_database.json)",
                DefaultValueFactory = _ => config?.Format ?? "standard"
            };
            formatOption.AcceptOnlyFromAmong("standard", "rich");

            var deduplicateOption = new Option<bool>("--deduplicate")
            {
                Description = "Merge duplicate entries for the same file into one best-compromise entry for IntelliSense",
                DefaultValueFactory = _ => config?.Deduplicate ?? false
            };

            var preferConfigOption = new Option<string>("--prefer-configuration")
            {
                Description = "Preferred configuration for conflict resolution during deduplication (default: Debug)",
                DefaultValueFactory = _ => config?.PreferConfiguration ?? "Debug"
            };

            var preferPlatformOption = new Option<string>("--prefer-platform")
            {
                Description = "Preferred platform for conflict resolution during deduplication (default: x64)",
                DefaultValueFactory = _ => config?.PreferPlatform ?? "x64"
            };

            var listInstancesOption = new Option<bool>("--list-instances", "--list-vs")
            {
                Description = "List all Visual Studio installations and exit"
            };

            var vsInstanceOption = new Option<string?>("--vs-instance")
            {
                Description = "Select VS installation by instance ID (use --list-instances to see available)",
                DefaultValueFactory = _ => config?.VsInstance
            };

            var cCppPropertiesOption = new Option<bool>("--c-cpp-properties")
            {
                Description = "Emit a .vscode/c_cpp_properties.json pointing to the generated compile_commands.json",
                DefaultValueFactory = _ => config?.EmitCCppProperties ?? false
            };

            var emitDefaultsOption = new Option<bool>("--emit-defaults")
            {
                Description = "Include the project-wide default compile entry in the output (synthetic __project_defaults.cpp with the baseline switches for the project)",
                DefaultValueFactory = _ => config?.EmitDefaults ?? false
            };

            var mergeDefaultsOption = new Option<bool>("--merge-defaults")
            {
                Description = "Merge project-wide default switches (defines, language standard, warning level, etc.) into each per-file entry when they are not already present",
                DefaultValueFactory = _ => config?.MergeDefaults ?? false
            };

            var msbuildPropertyOption = new Option<string[]>("--msbuild-property")
            {
                Description = "Pass an MSBuild global property as KEY=VALUE (repeatable). Overrides built-in defaults (e.g. BuildProjectReferences=true).",
                AllowMultipleArgumentsPerToken = false
            };

            var msbuildEnvOption = new Option<string[]>("--msbuild-env")
            {
                Description = "Set an environment variable for the MSBuild process as KEY=VALUE (repeatable). Overrides built-in defaults.",
                AllowMultipleArgumentsPerToken = false
            };

            var msbuildLauncherOption = new Option<string>("--msbuild-launcher")
            {
                Description = "How to launch MSBuild: auto (sniff extension, default), cmd (force cmd.exe /c wrapper), direct (run executable directly), dotnet (force dotnet exec)",
                DefaultValueFactory = _ => config?.MsBuildLauncher ?? "auto"
            };
            msbuildLauncherOption.AcceptOnlyFromAmong("auto", "cmd", "direct", "dotnet");

            var includePathOrderOption = new Option<string>("--include-path-order")
            {
                Description = "Where to place include paths from MSBuild's IncludePath/ExternalIncludePath properties: auto (per-path heuristic, default), prepend (before /I), append (after /I, matches cl.exe INCLUDE-env semantics)",
                DefaultValueFactory = _ => config?.IncludePathOrder ?? "auto"
            };
            includePathOrderOption.AcceptOnlyFromAmong("auto", "prepend", "append");

            var configFileOption = new Option<string?>("--config")
            {
                Description = "Path to a JSON config file supplying option defaults. When omitted, a 'msbuild-extractor.json' in the current directory is used automatically if present. Command-line options override values from the file."
            };

            var rootCommand = new RootCommand("Extract compile_commands.json from Visual C++ MSBuild projects");
            rootCommand.Options.Add(projectOption);
            rootCommand.Options.Add(solutionOption);
            rootCommand.Options.Add(configOption);
            rootCommand.Options.Add(platformOption);
            rootCommand.Options.Add(vsPathOption);
            rootCommand.Options.Add(vcTargetsPathOption);
            rootCommand.Options.Add(clPathOption);
            rootCommand.Options.Add(solutionDirOption);
            rootCommand.Options.Add(vcToolsInstallDirOption);
            rootCommand.Options.Add(msbuildPathOption);
            rootCommand.Options.Add(outputOption);
            rootCommand.Options.Add(loggerOption);
            rootCommand.Options.Add(useDevEnvOption);
            rootCommand.Options.Add(allConfigsOption);
            rootCommand.Options.Add(mergeOption);
            rootCommand.Options.Add(strictOption);
            rootCommand.Options.Add(validateOption);
            rootCommand.Options.Add(formatOption);
            rootCommand.Options.Add(deduplicateOption);
            rootCommand.Options.Add(preferConfigOption);
            rootCommand.Options.Add(preferPlatformOption);
            rootCommand.Options.Add(listInstancesOption);
            rootCommand.Options.Add(vsInstanceOption);
            rootCommand.Options.Add(cCppPropertiesOption);
            rootCommand.Options.Add(emitDefaultsOption);
            rootCommand.Options.Add(mergeDefaultsOption);
            rootCommand.Options.Add(msbuildPropertyOption);
            rootCommand.Options.Add(msbuildEnvOption);
            rootCommand.Options.Add(msbuildLauncherOption);
            rootCommand.Options.Add(includePathOrderOption);
            rootCommand.Options.Add(configFileOption);

            rootCommand.Validators.Add(commandResult =>
            {
                var listInstances = commandResult.GetValue(listInstancesOption);
                if (listInstances)
                    return;

                var projects = commandResult.GetValue(projectOption) ?? [];
                var solutions = commandResult.GetValue(solutionOption) ?? [];

                // Inputs may also come from the config file. A validator cannot see values produced
                // by an option's DefaultValueFactory, so read them from the config directly.
                var hasConfigInputs = (config?.Projects?.Length ?? 0) > 0 || (config?.Solutions?.Length ?? 0) > 0;

                if (projects.Length == 0 && solutions.Length == 0 && !hasConfigInputs)
                    commandResult.AddError("At least one --project or --solution must be specified (via the command line or a config file).");

                if (commandResult.GetValue(cCppPropertiesOption) && commandResult.GetValue(formatOption) == "rich")
                    commandResult.AddError("--c-cpp-properties cannot be used with --format rich (rich format does not produce compile_commands.json).");
            });

            rootCommand.SetAction(parseResult =>
            {
                result = new CommandLineOptions
                {
                    Projects = parseResult.GetValue(projectOption) ?? [],
                    Solutions = parseResult.GetValue(solutionOption) ?? [],
                    Configuration = parseResult.GetValue(configOption)!,
                    Platform = parseResult.GetValue(platformOption)!,
                    VsPath = parseResult.GetValue(vsPathOption),
                    VcTargetsPath = parseResult.GetValue(vcTargetsPathOption),
                    ClPath = parseResult.GetValue(clPathOption),
                    SolutionDir = parseResult.GetValue(solutionDirOption),
                    VcToolsInstallDir = parseResult.GetValue(vcToolsInstallDirOption),
                    MsBuildPath = parseResult.GetValue(msbuildPathOption),
                    Output = parseResult.GetValue(outputOption),
                    EnableLogger = parseResult.GetValue(loggerOption),
                    UseDevEnv = parseResult.GetValue(useDevEnvOption),
                    AllConfigurations = parseResult.GetValue(allConfigsOption),
                    Merge = parseResult.GetValue(mergeOption),
                    Strict = parseResult.GetValue(strictOption),
                    Validate = parseResult.GetValue(validateOption),
                    Format = parseResult.GetValue(formatOption)!,
                    Deduplicate = parseResult.GetValue(deduplicateOption),
                    PreferConfiguration = parseResult.GetValue(preferConfigOption)!,
                    PreferPlatform = parseResult.GetValue(preferPlatformOption)!,
                    ListInstances = parseResult.GetValue(listInstancesOption),
                    VsInstance = parseResult.GetValue(vsInstanceOption),
                    EmitCCppProperties = parseResult.GetValue(cCppPropertiesOption),
                    EmitDefaults = parseResult.GetValue(emitDefaultsOption),
                    MergeDefaults = parseResult.GetValue(mergeDefaultsOption),
                    MsBuildProperties = MergeKeyValues(config?.MsBuildProperties,
                        ParseKeyValuePairs(parseResult.GetValue(msbuildPropertyOption), "--msbuild-property")),
                    MsBuildEnv = MergeKeyValues(config?.MsBuildEnv,
                        ParseKeyValuePairs(parseResult.GetValue(msbuildEnvOption), "--msbuild-env")),
                    MsBuildLauncher = parseResult.GetValue(msbuildLauncherOption)!,
                    IncludePathOrder = parseResult.GetValue(includePathOrderOption)!,
                    ConfigPath = config?.SourcePath
                };

                // Re-check option combinations against the effective (merged) values. As with the
                // input check above, a validator cannot see config-supplied defaults, so a
                // config-only "rich + c-cpp-properties" clash would otherwise slip through.
                if (result.EmitCCppProperties && result.Format == "rich")
                {
                    Console.Error.WriteLine("Error: --c-cpp-properties / \"emitCCppProperties\" cannot be used with the 'rich' format (rich format does not produce compile_commands.json).");
                    Environment.Exit(1);
                }
            });

            var exitCode = rootCommand.Parse(args).Invoke();

            if (result == null)
                Environment.Exit(exitCode);

            return result;
        }

        private static Dictionary<string, string> ParseKeyValuePairs(string[]? values, string flagName)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (values == null) return dict;
            foreach (var entry in values)
            {
                var separatorIndex = entry.IndexOf('=');
                if (separatorIndex <= 0)
                    throw new ArgumentException($"{flagName} value '{entry}' must be in KEY=VALUE form.");
                dict[entry.Substring(0, separatorIndex)] = entry.Substring(separatorIndex + 1);
            }
            return dict;
        }

        private static Dictionary<string, string> MergeKeyValues(
            Dictionary<string, string>? fromConfig, Dictionary<string, string> fromCommandLine)
        {
            var merged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (fromConfig != null)
                foreach (var kvp in fromConfig)
                    merged[kvp.Key] = kvp.Value;
            foreach (var kvp in fromCommandLine)
                merged[kvp.Key] = kvp.Value;
            return merged;
        }
    }
}
