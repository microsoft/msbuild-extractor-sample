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
                return null; // Cannot validate, no configs found in project

            var match = configs.Any(cp =>
                cp.Configuration.Equals(configuration, StringComparison.OrdinalIgnoreCase) &&
                cp.Platform.Equals(platform, StringComparison.OrdinalIgnoreCase));

            if (match)
                return null;

            var available = string.Join(", ", configs.Select(cp => cp.ToString()));
            return $"Configuration '{configuration}|{platform}' not found in {Path.GetFileName(projectPath)}. Available: {available}";
        }

        /// <summary>
        /// Discovers VC++ projects referenced by an MSBuild "dirs.proj" traversal
        /// project (Sdk="Microsoft.Build.Traversal"), recursing into nested traversal
        /// projects. Non-VC references (.csproj, .props, etc.) are skipped. Mirrors
        /// <see cref="GetVcProjectsFromSolution"/> for .sln/.slnx inputs.
        /// </summary>
        public static List<VcProject> GetVcProjectsFromDirsProj(string dirsProjPath)
        {
            dirsProjPath = Path.GetFullPath(dirsProjPath);

            if (!File.Exists(dirsProjPath))
                throw new FileNotFoundException($"dirs project file not found: {dirsProjPath}");

            var projects = new List<VcProject>();
            var visitedTraversals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenVcxproj = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectVcProjectsFromTraversal(dirsProjPath, projects, visitedTraversals, seenVcxproj);
            return projects;
        }

        public static List<string> GetVcProjectPathsFromDirsProj(string dirsProjPath) =>
            GetVcProjectsFromDirsProj(dirsProjPath).Select(p => p.Path).ToList();

        private static void CollectVcProjectsFromTraversal(
            string traversalPath,
            List<VcProject> projects,
            HashSet<string> visitedTraversals,
            HashSet<string> seenVcxproj)
        {
            traversalPath = Path.GetFullPath(traversalPath);

            // Guard against cycles and repeated traversal of the same dirs.proj.
            if (!visitedTraversals.Add(traversalPath))
                return;

            if (!File.Exists(traversalPath))
            {
                Console.Error.WriteLine($"Warning: referenced traversal project not found: {traversalPath}");
                return;
            }

            XDocument doc;
            try
            {
                doc = XDocument.Load(traversalPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: failed to parse traversal project {traversalPath}: {ex.Message}");
                return;
            }

            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
            var baseDir = Path.GetDirectoryName(traversalPath)!;

            // Microsoft.Build.Traversal uses <ProjectReference Include="..."/>. Also accept
            // <ProjectFile Include="..."/> used by some legacy traversal projects.
            var includes = doc.Descendants(ns + "ProjectReference")
                .Concat(doc.Descendants(ns + "ProjectFile"))
                .Select(e => e.Attribute("Include")?.Value)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Cast<string>();

            foreach (var include in includes)
            {
                foreach (var resolved in ResolveTraversalInclude(baseDir, include))
                {
                    var ext = Path.GetExtension(resolved);

                    if (ext.Equals(".vcxproj", StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(resolved) && seenVcxproj.Add(resolved))
                        {
                            projects.Add(new VcProject(
                                Path: resolved,
                                Name: Path.GetFileNameWithoutExtension(resolved),
                                ConfigurationPlatforms: GetProjectConfigurations(resolved)));
                        }
                    }
                    else if (ext.Equals(".proj", StringComparison.OrdinalIgnoreCase))
                    {
                        // Nested traversal project (e.g. a subdirectory dirs.proj). Recurse.
                        CollectVcProjectsFromTraversal(resolved, projects, visitedTraversals, seenVcxproj);
                    }
                    // Other project types (.csproj, .props, etc.) are not C/C++ and are skipped.
                }
            }
        }

        // Resolves a traversal ProjectReference Include value into concrete file paths.
        // Handles ';'-separated lists and simple '*'/'**' wildcard globs. Entries that
        // contain unresolved MSBuild property/item expansions ($(..)/%(..)) are skipped.
        private static IEnumerable<string> ResolveTraversalInclude(string baseDir, string include)
        {
            foreach (var raw in include.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (raw.Contains("$(") || raw.Contains("%("))
                    continue;

                var combined = Path.IsPathRooted(raw) ? raw : Path.Combine(baseDir, raw);

                if (raw.Contains('*'))
                {
                    foreach (var match in ExpandWildcard(combined))
                        yield return Path.GetFullPath(match);
                }
                else
                {
                    yield return Path.GetFullPath(combined);
                }
            }
        }

        // Minimal wildcard expander for traversal Include globs. Supports a literal
        // directory prefix followed by '*'/'**' and a trailing filename pattern
        // (e.g. "src\\**\\*.vcxproj"). Best-effort: unmatched or malformed patterns yield nothing.
        private static IEnumerable<string> ExpandWildcard(string pattern)
        {
            var separators = new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar };
            var firstWildcard = pattern.IndexOf('*');
            if (firstWildcard < 0)
                yield break;

            var lastSep = pattern.LastIndexOfAny(separators, firstWildcard);
            if (lastSep < 0)
                yield break;

            var globBaseDir = pattern.Substring(0, lastSep);
            var remainder = pattern.Substring(lastSep + 1);
            if (!Directory.Exists(globBaseDir))
                yield break;

            var recursive = remainder.Contains("**");
            var searchPattern = Path.GetFileName(remainder);
            if (string.IsNullOrEmpty(searchPattern) || searchPattern == "**")
                searchPattern = "*";

            List<string> files;
            try
            {
                files = Directory.EnumerateFiles(
                    globBaseDir, searchPattern,
                    recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly).ToList();
            }
            catch
            {
                yield break;
            }

            foreach (var f in files)
                yield return f;
        }
    }
}
