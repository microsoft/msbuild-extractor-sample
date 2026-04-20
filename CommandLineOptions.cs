using System.CommandLine;

namespace MSBuild.CompileCommands.Extractor
{
    public class CommandLineOptions
    {
        public string[] Projects { get; set; } = [];
        public string[] Solutions { get; set; } = [];

        // Convenience accessors for single-input backward compatibility
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

        public static CommandLineOptions Parse(string[] args)
        {
            CommandLineOptions? result = null;

            var projectOption = new Option<string[]>("--project", "-p")
            {
                Description = "Path to .vcxproj file (can be specified multiple times)",
                AllowMultipleArgumentsPerToken = true
            };

            var solutionOption = new Option<string[]>("--solution", "-s")
            {
                Description = "Path to .sln or .slnx file (can be specified multiple times)",
                AllowMultipleArgumentsPerToken = true
            };

            var configOption = new Option<string>("--configuration", "-c")
            {
                Description = "Build configuration",
                DefaultValueFactory = _ => "Debug"
            };

            var platformOption = new Option<string>("--platform", "-a")
            {
                Description = "Build platform",
                DefaultValueFactory = _ => "x64"
            };

            var vsPathOption = new Option<string?>("--vs-path")
            {
                Description = "Path to Visual Studio installation"
            };

            var vcTargetsPathOption = new Option<string?>("--vc-targets-path")
            {
                Description = "Path to VC targets (e.g. ...\\MSBuild\\Microsoft\\VC\\v180)"
            };

            var clPathOption = new Option<string?>("--cl-path")
            {
                Description = "Path to cl.exe"
            };

            var solutionDirOption = new Option<string?>("--solution-dir")
            {
                Description = "Value for the SolutionDir MSBuild property (auto-derived when using --solution)"
            };

            var vcToolsInstallDirOption = new Option<string?>("--vc-tools-install-dir")
            {
                Description = "Value for the VCToolsInstallDir MSBuild property"
            };

            var outputOption = new Option<string?>("--output", "-o")
            {
                Description = "Output path for compile_commands.json"
            };

            var loggerOption = new Option<bool>("--logger")
            {
                Description = "Enable MSBuild console logger output",
                DefaultValueFactory = _ => false
            };

            var useDevEnvOption = new Option<bool>("--use-dev-env")
            {
                Description = "Read environment variables from Developer Command Prompt (VCToolsInstallDir, VCTargetsPath, etc.)",
                DefaultValueFactory = _ => false
            };

            var msbuildPathOption = new Option<string?>("--msbuild-path")
            {
                Description = "Path to msbuild.exe (enables out-of-process mode for custom MSBuild/toolchain locations)"
            };

            var allConfigsOption = new Option<bool>("--all-configurations")
            {
                Description = "Extract for all configuration/platform combinations found in the project or solution",
                DefaultValueFactory = _ => false
            };

            var mergeOption = new Option<bool>("--merge")
            {
                Description = "Merge all configurations into a single output file (use with --all-configurations)",
                DefaultValueFactory = _ => false
            };

            var strictOption = new Option<bool>("--strict")
            {
                Description = "Treat configuration/platform validation warnings as errors",
                DefaultValueFactory = _ => false
            };

            var validateOption = new Option<bool>("--validate")
            {
                Description = "After extraction, verify each compile command by running cl.exe /c",
                DefaultValueFactory = _ => false
            };

            var formatOption = new Option<string>("--format", "-f")
            {
                Description = "Output format: 'standard' (compile_commands.json) or 'rich' (hierarchical compile_database.json)",
                DefaultValueFactory = _ => "standard"
            };
            formatOption.AcceptOnlyFromAmong("standard", "rich");

            var deduplicateOption = new Option<bool>("--deduplicate")
            {
                Description = "Merge duplicate entries for the same file into one best-compromise entry for IntelliSense",
                DefaultValueFactory = _ => false
            };

            var preferConfigOption = new Option<string>("--prefer-configuration")
            {
                Description = "Preferred configuration for conflict resolution during deduplication (default: Debug)",
                DefaultValueFactory = _ => "Debug"
            };

            var preferPlatformOption = new Option<string>("--prefer-platform")
            {
                Description = "Preferred platform for conflict resolution during deduplication (default: x64)",
                DefaultValueFactory = _ => "x64"
            };

            var listInstancesOption = new Option<bool>("--list-instances", "--list-vs")
            {
                Description = "List all Visual Studio installations and exit"
            };

            var vsInstanceOption = new Option<string?>("--vs-instance")
            {
                Description = "Select VS installation by instance ID (use --list-instances to see available)"
            };

            var cCppPropertiesOption = new Option<bool>("--c-cpp-properties")
            {
                Description = "Emit a .vscode/c_cpp_properties.json pointing to the generated compile_commands.json",
                DefaultValueFactory = _ => false
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

            rootCommand.Validators.Add(commandResult =>
            {
                var listInstances = commandResult.GetValue(listInstancesOption);
                if (listInstances)
                    return;

                var projects = commandResult.GetValue(projectOption) ?? [];
                var solutions = commandResult.GetValue(solutionOption) ?? [];

                if (projects.Length == 0 && solutions.Length == 0)
                    commandResult.AddError("At least one --project or --solution must be specified.");

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
                    EmitCCppProperties = parseResult.GetValue(cCppPropertiesOption)
                };
            });

            var exitCode = rootCommand.Parse(args).Invoke();

            if (result == null)
                Environment.Exit(exitCode);

            return result;
        }
    }
}
