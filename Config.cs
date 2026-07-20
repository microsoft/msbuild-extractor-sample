using System.Text.Json;

namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Options loaded from an optional JSON config file. Every member is nullable so an absent key
    /// leaves the matching command-line default untouched. Keys bind case-insensitively, and the
    /// loader tolerates // comments and trailing commas.
    /// </summary>
    public sealed class Config
    {
        /// <summary>File name looked up in the current directory when --config is not passed.</summary>
        public const string DefaultFileName = "msbuild-extractor.json";

        public string[]? Projects { get; set; }
        public string[]? Solutions { get; set; }
        public string? Configuration { get; set; }
        public string? Platform { get; set; }
        public string? Output { get; set; }
        public string? Format { get; set; }

        public bool? AllConfigurations { get; set; }
        public bool? Merge { get; set; }
        public bool? Deduplicate { get; set; }
        public string? PreferConfiguration { get; set; }
        public string? PreferPlatform { get; set; }

        public bool? Strict { get; set; }
        public bool? Validate { get; set; }
        public bool? Logger { get; set; }
        public bool? EmitCCppProperties { get; set; }
        public bool? EmitDefaults { get; set; }
        public bool? MergeDefaults { get; set; }

        public bool? UseDevEnv { get; set; }
        public string? VsInstance { get; set; }
        public string? VsPath { get; set; }
        public string? MsBuildPath { get; set; }
        public string? VcTargetsPath { get; set; }
        public string? ClPath { get; set; }
        public string? VcToolsInstallDir { get; set; }
        public string? SolutionDir { get; set; }

        public string? MsBuildLauncher { get; set; }
        public string? IncludePathOrder { get; set; }

        public Dictionary<string, string>? MsBuildProperties { get; set; }
        public Dictionary<string, string>? MsBuildEnv { get; set; }

        /// <summary>The file this config was loaded from, or null when no config was applied.</summary>
        public string? SourcePath { get; private set; }

        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

        /// <summary>
        /// Finds a config file (an explicit --config &lt;path&gt;, otherwise a msbuild-extractor.json in
        /// the current directory) and loads it into this instance. Returns true when one was applied.
        /// Exits with an error on a missing explicit path, malformed JSON, or an invalid value.
        /// </summary>
        public bool Load(string[] args)
        {
            var path = FindPath(args);
            if (path == null)
                return false;

            Config data;
            try
            {
                data = JsonSerializer.Deserialize<Config>(File.ReadAllText(path), SerializerOptions) ?? new Config();
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"Error: invalid config file '{path}': {ex.Message}");
                Environment.Exit(1);
                return false;
            }

            Projects = data.Projects;
            Solutions = data.Solutions;
            Configuration = data.Configuration;
            Platform = data.Platform;
            Output = data.Output;
            Format = data.Format;
            AllConfigurations = data.AllConfigurations;
            Merge = data.Merge;
            Deduplicate = data.Deduplicate;
            PreferConfiguration = data.PreferConfiguration;
            PreferPlatform = data.PreferPlatform;
            Strict = data.Strict;
            Validate = data.Validate;
            Logger = data.Logger;
            EmitCCppProperties = data.EmitCCppProperties;
            EmitDefaults = data.EmitDefaults;
            MergeDefaults = data.MergeDefaults;
            UseDevEnv = data.UseDevEnv;
            VsInstance = data.VsInstance;
            VsPath = data.VsPath;
            MsBuildPath = data.MsBuildPath;
            VcTargetsPath = data.VcTargetsPath;
            ClPath = data.ClPath;
            VcToolsInstallDir = data.VcToolsInstallDir;
            SolutionDir = data.SolutionDir;
            MsBuildLauncher = data.MsBuildLauncher;
            IncludePathOrder = data.IncludePathOrder;
            MsBuildProperties = data.MsBuildProperties;
            MsBuildEnv = data.MsBuildEnv;

            SourcePath = path;
            ValidateValues(path);
            ResolvePaths(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
            return true;
        }

        private string? FindPath(string[] args)
        {
            var explicitPath = GetConfigArgument(args);
            if (explicitPath != null)
            {
                if (!File.Exists(explicitPath))
                {
                    Console.Error.WriteLine($"Error: config file not found: {explicitPath}");
                    Environment.Exit(1);
                }
                return Path.GetFullPath(explicitPath);
            }

            var auto = Path.Combine(Directory.GetCurrentDirectory(), DefaultFileName);
            return File.Exists(auto) ? auto : null;
        }

        private string? GetConfigArgument(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                if (arg == "--config")
                    return i + 1 < args.Length ? args[i + 1] : null;
                if (arg.StartsWith("--config=", StringComparison.Ordinal))
                    return arg.Substring("--config=".Length);
            }
            return null;
        }

        private void ValidateValues(string path)
        {
            void Check(string? value, string field, params string[] allowed)
            {
                if (value != null && !allowed.Contains(value, StringComparer.OrdinalIgnoreCase))
                {
                    Console.Error.WriteLine(
                        $"Error: config file '{path}': '{field}' must be one of [{string.Join(", ", allowed)}] but was '{value}'.");
                    Environment.Exit(1);
                }
            }

            Check(Format, "format", "standard", "rich");
            Check(MsBuildLauncher, "msBuildLauncher", "auto", "cmd", "direct", "dotnet");
            Check(IncludePathOrder, "includePathOrder", "auto", "prepend", "append");
        }

        private void ResolvePaths(string baseDir)
        {
            Projects = ResolveEach(Projects, baseDir, exe: false);
            Solutions = ResolveEach(Solutions, baseDir, exe: false);
            Output = ResolveOne(Output, baseDir, exe: false);
            VsPath = ResolveOne(VsPath, baseDir, exe: false);
            VcTargetsPath = ResolveOne(VcTargetsPath, baseDir, exe: false);
            VcToolsInstallDir = ResolveOne(VcToolsInstallDir, baseDir, exe: false);
            SolutionDir = ResolveOne(SolutionDir, baseDir, exe: false);
            MsBuildPath = ResolveOne(MsBuildPath, baseDir, exe: true);
            ClPath = ResolveOne(ClPath, baseDir, exe: true);
        }

        private string[]? ResolveEach(string[]? values, string baseDir, bool exe)
            => values?.Select(v => ResolveOne(v, baseDir, exe)!).ToArray();

        private string? ResolveOne(string? value, string baseDir, bool exe)
        {
            if (string.IsNullOrEmpty(value) || Path.IsPathRooted(value))
                return value;
            // Bare executable names (no directory separator) are left alone so PATH lookup still works.
            if (exe && !value.Contains('/') && !value.Contains('\\'))
                return value;
            return Path.GetFullPath(Path.Combine(baseDir, value));
        }
    }
}
