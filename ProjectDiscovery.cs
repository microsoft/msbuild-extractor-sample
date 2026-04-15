using Microsoft.Build.Construction;
using System.Xml.Linq;

namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Discovers VC++ projects from solutions and reads project configurations.
    /// Supports .sln, .slnx, and .vcxproj files.
    /// </summary>
    public static class ProjectDiscovery
    {
        public static List<VcProject> GetVcProjectsFromSolution(string solutionPath)
        {
            solutionPath = Path.GetFullPath(solutionPath);

            if (!File.Exists(solutionPath))
                throw new FileNotFoundException($"Solution file not found: {solutionPath}");

            if (solutionPath.EndsWith(".slnx", StringComparison.OrdinalIgnoreCase))
                return GetVcProjectsFromSlnx(solutionPath);
            else
                return GetVcProjectsFromSln(solutionPath);
        }

        public static List<string> GetVcProjectPathsFromSolution(string solutionPath) =>
            GetVcProjectsFromSolution(solutionPath).Select(p => p.Path).ToList();

        private static List<VcProject> GetVcProjectsFromSln(string solutionPath)
        {
            var solutionFile = SolutionFile.Parse(solutionPath);
            var vcxProjects = new List<VcProject>();

            foreach (var project in solutionFile.ProjectsInOrder)
            {
                if (project.ProjectType == SolutionProjectType.KnownToBeMSBuildFormat &&
                    project.AbsolutePath.EndsWith(".vcxproj", StringComparison.OrdinalIgnoreCase))
                {
                    var configPlatforms = project.ProjectConfigurations.Values
                        .Select(c => new ConfigurationPlatform(c.ConfigurationName, c.PlatformName))
                        .Distinct()
                        .ToArray();

                    vcxProjects.Add(new VcProject(
                        Path: project.AbsolutePath,
                        Name: project.ProjectName,
                        ConfigurationPlatforms: configPlatforms
                    ));
                }
            }

            return vcxProjects;
        }

        private static List<VcProject> GetVcProjectsFromSlnx(string solutionPath)
        {
            var solutionDir = Path.GetDirectoryName(solutionPath)!;
            var vcxProjects = new List<VcProject>();

            var doc = XDocument.Load(solutionPath);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            var projects = doc.Descendants(ns + "Project")
                .Where(p => p.Attribute("Path") != null);

            foreach (var project in projects)
            {
                var relativePath = project.Attribute("Path")!.Value;

                if (!relativePath.EndsWith(".vcxproj", StringComparison.OrdinalIgnoreCase))
                    continue;

                var absolutePath = Path.GetFullPath(Path.Combine(solutionDir, relativePath));
                var projectName = Path.GetFileNameWithoutExtension(relativePath);

                var buildTypes = doc.Descendants(ns + "BuildType")
                    .Select(b => b.Attribute("Name")?.Value)
                    .Where(v => v != null)
                    .Cast<string>()
                    .ToList();

                var platformElements = doc.Descendants(ns + "Platform")
                    .Select(p => p.Attribute("Name")?.Value)
                    .Where(v => v != null)
                    .Cast<string>()
                    .ToList();

                if (buildTypes.Count == 0)
                {
                    buildTypes.Add("Debug");
                    buildTypes.Add("Release");
                }
                if (platformElements.Count == 0)
                {
                    platformElements.Add("x64");
                }

                // .slnx doesn't have per-project config mappings; create cross-product
                var configPlatforms = buildTypes
                    .SelectMany(c => platformElements.Select(p => new ConfigurationPlatform(c, p)))
                    .Distinct()
                    .ToArray();

                vcxProjects.Add(new VcProject(
                    Path: absolutePath,
                    Name: projectName,
                    ConfigurationPlatforms: configPlatforms
                ));
            }

            return vcxProjects;
        }

        /// <summary>
        /// Reads ProjectConfiguration items from a .vcxproj file to discover
        /// available configuration/platform pairs.
        /// </summary>
        public static ConfigurationPlatform[] GetProjectConfigurations(string vcxprojPath)
        {
            try
            {
                var doc = XDocument.Load(vcxprojPath);
                var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

                return doc.Descendants(ns + "ProjectConfiguration")
                    .Select(pc =>
                    {
                        // Try Include="Debug|x64" format first
                        var include = pc.Attribute("Include")?.Value;
                        if (include != null && include.Contains('|'))
                        {
                            var parts = include.Split('|', 2);
                            return new ConfigurationPlatform(parts[0], parts[1]);
                        }
                        // Fallback: read child elements
                        var config = pc.Element(ns + "Configuration")?.Value;
                        var platform = pc.Element(ns + "Platform")?.Value;
                        if (config != null && platform != null)
                            return new ConfigurationPlatform(config, platform);
                        return null;
                    })
                    .Where(cp => cp != null)
                    .Cast<ConfigurationPlatform>()
                    .Distinct()
                    .ToArray();
            }
            catch
            {
                return [];
            }
        }

        /// <summary>
        /// Validates that the requested configuration/platform exists.
        /// Returns null if valid, or an error message with available combos if invalid.
        /// </summary>
        public static string? ValidateConfigurationPlatform(
            string projectPath,
            string configuration,
            string platform,
            ConfigurationPlatform[]? knownConfigs = null)
        {
            var configs = knownConfigs ?? GetProjectConfigurations(projectPath);
            if (configs.Length == 0)
                return null; // Cannot validate — no configs found in project

            var match = configs.Any(cp =>
                cp.Configuration.Equals(configuration, StringComparison.OrdinalIgnoreCase) &&
                cp.Platform.Equals(platform, StringComparison.OrdinalIgnoreCase));

            if (match)
                return null;

            var available = string.Join(", ", configs.Select(cp => cp.ToString()));
            return $"Configuration '{configuration}|{platform}' not found in {Path.GetFileName(projectPath)}. Available: {available}";
        }
    }
}
