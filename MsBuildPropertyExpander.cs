using Microsoft.Build.Execution;
using System.Text.RegularExpressions;

namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Expands residual MSBuild property references (e.g. <c>$(OpenConsoleDir)</c>,
    /// <c>$(GeneratedFilesDir)</c>) that survive into extracted compile commands, using a project's
    /// evaluated property values. MSBuild sometimes captures a command line, file path, or directory
    /// before every <c>$(...)</c> in metadata like AdditionalIncludeDirectories has been expanded, so
    /// this pass substitutes the authoritative evaluated value the project already computed.
    /// Shared by both the in-process and out-of-process extractors.
    /// </summary>
    public static class MsBuildPropertyExpander
    {
        // Matches a single MSBuild property reference, e.g. $(OpenConsoleDir).
        private static readonly Regex MsBuildPropertyRef = new(@"\$\(([^)]+)\)", RegexOptions.Compiled);

        private const int MaxPropertyExpansionDepth = 10;

        /// <summary>
        /// Expands residual MSBuild property references in each element of an argument array.
        /// Expansion is case-insensitive and re-scans its own output (bounded by
        /// <see cref="MaxPropertyExpansionDepth"/>) so nested references resolve too.
        /// Returns the input unchanged when no properties are available.
        /// </summary>
        public static string[] ExpandMsBuildProperties(string[] args, IReadOnlyDictionary<string, string>? evaluatedProperties)
        {
            if (evaluatedProperties == null || evaluatedProperties.Count == 0)
                return args;

            var result = new string[args.Length];
            for (int i = 0; i < args.Length; i++)
                result[i] = ExpandOne(args[i], evaluatedProperties);
            return result;
        }

        /// <summary>
        /// Expands residual MSBuild property references in a single string (e.g. a source file
        /// path or working directory). Returns the input unchanged when no properties are available.
        /// </summary>
        public static string ExpandMsBuildProperties(string value, IReadOnlyDictionary<string, string>? evaluatedProperties)
        {
            if (string.IsNullOrEmpty(value) || evaluatedProperties == null || evaluatedProperties.Count == 0)
                return value;
            return ExpandOne(value, evaluatedProperties);
        }

        private static string ExpandOne(string value, IReadOnlyDictionary<string, string> evaluatedProperties)
        {
            for (int depth = 0; depth < MaxPropertyExpansionDepth && value.Contains("$("); depth++)
            {
                bool replacedAny = false;
                value = MsBuildPropertyRef.Replace(value, match =>
                {
                    var name = match.Groups[1].Value;
                    // Substitute whenever the property is DEFINED, even if it evaluates to an empty
                    // string. A defined-but-empty property (common at design time for values that
                    // depend on build-time state, e.g. $(GeneratedFilesDir) deriving from $(IntDir))
                    // should still have its $(...) removed rather than leaking into the output.
                    if (evaluatedProperties.TryGetValue(name, out var resolved))
                    {
                        replacedAny = true;
                        return resolved ?? string.Empty;
                    }

                    // Unresolved residual: the property is entirely absent from the evaluated
                    // project properties, so we intentionally leave the literal $(...) untouched
                    // here rather than guessing or dropping the argument. If you are chasing an
                    // unexpanded variable in compile_commands.json (e.g. $(OpenConsoleDir) or
                    // $(GeneratedFilesDir)), this is the spot: the property name did not appear in
                    // the evaluated project properties passed in (post-build ProjectInstance state
                    // in-process / binlog Property nodes out-of-process).
                    return match.Value;
                });

                if (!replacedAny)
                    break;
            }

            return value;
        }

        /// <summary>
        /// Builds a case-insensitive map of evaluated MSBuild property name to value from a
        /// <see cref="ProjectInstance"/>, used to expand residual $(...) references (in-process).
        /// </summary>
        public static Dictionary<string, string> SnapshotProperties(ProjectInstance projectInstance)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var prop in projectInstance.Properties)
            {
                if (!string.IsNullOrEmpty(prop.Name))
                    map[prop.Name] = prop.EvaluatedValue ?? string.Empty;
            }
            return map;
        }
    }
}
