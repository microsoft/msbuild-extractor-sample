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
        string[] Files
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
