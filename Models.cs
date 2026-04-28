using System.Text.Json.Serialization;

namespace MSBuild.CompileCommands.Extractor
{
    public record ConfigurationPlatform(string Configuration, string Platform)
    {
        public override string ToString() => $"{Configuration}|{Platform}";
    }

    public record ClCommandLineEntry(
        string CommandLine,
        string WorkingDirectory,
        string ToolPath,
        string[] Files,
        // True when this entry came from the synthetic __temporary.cpp probe injected by
        // InitGetClCommandLines in Microsoft.Cpp.DesignTime.targets (tag: ConfigurationOptions=true).
        // Represents the project-wide default switches used for files with no per-file overrides.
        bool IsConfigurationDefault = false,
        // True when the entry originated from @(FxCompile) (HLSL) rather than @(ClCompile).
        // The Microsoft.Cpp.DesignTime.targets target runs CLCommandLine separately for
        // FxCompile and merges output into the same @(ClCommandLines) group. FxCompile
        // switches are fxc.exe flags, not cl.exe flags, and are not consumable by clangd,
        // so we drop them from standard compile_commands.json output.
        bool IsFxCompile = false
    );

    /// <summary>
    /// Project-level IntelliSense directories extracted from the GetProjectDirectories
    /// target (one _ProjectDirectories item).
    /// </summary>
    public record ProjectDirectories(
        string IncludePath,
        string ExternalIncludePath,
        string FrameworkIncludePath,
        string ExcludePath,
        string ReferencePath,
        string ProjectDir,
        string ToolsetISenseIdentifier
    );

    public interface ICompileCommandsExtractor
    {
        List<CompileCommand> ExtractCompileCommands();
        void WriteCompileCommandsJson(string? outputPath = null);
    }

    public record VcProject(
        string Path,
        string Name,
        ConfigurationPlatform[] ConfigurationPlatforms
    )
    {
        public string[] Configurations => ConfigurationPlatforms
            .Select(cp => cp.Configuration)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        public string[] Platforms => ConfigurationPlatforms
            .Select(cp => cp.Platform)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        public bool HasConfigurationPlatform(string configuration, string platform) =>
            ConfigurationPlatforms.Any(cp =>
                cp.Configuration.Equals(configuration, StringComparison.OrdinalIgnoreCase) &&
                cp.Platform.Equals(platform, StringComparison.OrdinalIgnoreCase));
    }

    public record CompileCommand(
        string File,
        string[] Arguments,
        string Directory,
        [property: JsonIgnore]
        string? ProjectPath = null,
        [property: JsonIgnore]
        string? ProjectName = null,
        [property: JsonIgnore]
        string? Configuration = null,
        [property: JsonIgnore]
        string? Platform = null
    );
}
