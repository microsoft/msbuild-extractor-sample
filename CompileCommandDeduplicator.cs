namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Merges duplicate compile command entries (same source file) into a single
    /// best-compromise entry suitable for IntelliSense tools like clangd/cpptools.
    /// </summary>
    public static class CompileCommandDeduplicator
    {
        private static readonly StringComparer PathComparer = StringComparer.OrdinalIgnoreCase;

        // Language standard ranking, higher index = newer standard
        private static readonly string[] CppStandardRank = ["c++14", "c++17", "c++20", "c++23", "c++latest"];
        private static readonly string[] CStandardRank = ["c11", "c17", "clatest"];

        /// <summary>
        /// Deduplicate compile commands so there is at most one entry per source file.
        /// </summary>
        public static List<CompileCommand> Deduplicate(List<CompileCommand> commands, string preferConfig = "Debug", string preferPlatform = "x64")
        {
            var groups = commands.GroupBy(c => c.File, PathComparer).ToList();

            var result = new List<CompileCommand>(groups.Count);

            foreach (var group in groups)
            {
                var entries = group.ToList();
                if (entries.Count == 1)
                {
                    // Single entry, still strip warnings/optimization flags
                    var cleaned = CleanSingleEntry(entries[0]);
                    result.Add(cleaned);
                }
                else
                {
                    result.Add(MergeEntries(entries, preferConfig, preferPlatform));
                }
            }

            return result;
        }

        private static CompileCommand CleanSingleEntry(CompileCommand entry)
        {
            var args = entry.Arguments.ToList();
            args = StripWarningFlags(args);
            args = StripOptimizationFlags(args);
            return entry with { Arguments = args.ToArray() };
        }

        private static CompileCommand MergeEntries(List<CompileCommand> entries, string preferConfig, string preferPlatform)
        {
            // Sort entries so preferred config/platform comes first
            var sorted = entries
                .OrderByDescending(e => IsPreferred(e, preferConfig, preferPlatform))
                .ToList();

            var preferred = sorted[0];

            // Parse arguments from all entries
            var parsedAll = sorted.Select(ParseArguments).ToList();

            // The cl.exe path and source file come from the preferred entry;
            // the source file is identical across every entry in the group,
            // since that's the grouping key.
            string clExe = parsedAll[0].ClExe;
            string sourceFile = parsedAll[0].SourceFile ?? preferred.File;

            // Include paths are unioned with case-insensitive, order-preserving dedup.
            var includes = UnionOrderPreserving(parsedAll.SelectMany(p => p.IncludePaths), PathComparer);

            // Defines need a smart merge to handle conflicting values across entries.
            var defines = MergeDefines(parsedAll.Select(p => p.Defines).ToList());

            // For language standard, keep the highest seen across the group.
            string? langStd = PickHighestStandard(parsedAll.Select(p => p.LanguageStandard).Where(s => s != null).Cast<string>());

            // For the runtime library, prefer the debug variant when present.
            string? runtimeLib = PickDebugRuntime(parsedAll.Select(p => p.RuntimeLibrary).Where(s => s != null).Cast<string>());

            // Remaining flags are unioned permissively.
            var otherFlags = UnionOrderPreserving(parsedAll.SelectMany(p => p.OtherFlags), StringComparer.OrdinalIgnoreCase);

            // The working directory comes from the preferred entry.
            string directory = preferred.Directory;

            // Reconstruct arguments
            var args = new List<string>();
            args.Add(clExe);

            foreach (var inc in includes)
                args.Add($"/I{inc}");

            foreach (var def in defines)
                args.Add($"/D{def}");

            if (langStd != null)
                args.Add($"/std:{langStd}");

            if (runtimeLib != null)
                args.Add($"/{runtimeLib}");

            args.AddRange(otherFlags);
            args.Add(sourceFile);

            return new CompileCommand(
                File: preferred.File,
                Arguments: args.ToArray(),
                Directory: directory,
                ProjectPath: preferred.ProjectPath,
                ProjectName: preferred.ProjectName,
                Configuration: preferred.Configuration,
                Platform: preferred.Platform);
        }

        private static int IsPreferred(CompileCommand entry, string preferConfig, string preferPlatform)
        {
            int score = 0;
            if (entry.Configuration != null &&
                entry.Configuration.Equals(preferConfig, StringComparison.OrdinalIgnoreCase))
                score += 2;
            if (entry.Platform != null &&
                entry.Platform.Equals(preferPlatform, StringComparison.OrdinalIgnoreCase))
                score += 1;
            return score;
        }

        #region Argument Parsing

        private record ParsedArgs(
            string ClExe,
            string? SourceFile,
            List<string> IncludePaths,
            List<string> Defines,
            string? LanguageStandard,
            string? RuntimeLibrary,
            List<string> OtherFlags);

        private static ParsedArgs ParseArguments(CompileCommand cmd)
        {
            var args = cmd.Arguments;
            string clExe = args.Length > 0 ? args[0] : "cl.exe";
            string? sourceFile = null;
            var includes = new List<string>();
            var defines = new List<string>();
            string? langStd = null;
            string? runtimeLib = null;
            var other = new List<string>();

            for (int i = 1; i < args.Length; i++)
            {
                var arg = args[i];

                if (arg.StartsWith("/external:I", StringComparison.OrdinalIgnoreCase) ||
                    arg.StartsWith("-external:I", StringComparison.OrdinalIgnoreCase))
                {
                    var path = arg.Substring(11); // "/external:I".Length
                    if (string.IsNullOrEmpty(path) && i + 1 < args.Length)
                        path = args[++i];
                    if (!string.IsNullOrEmpty(path))
                        includes.Add(path);
                }
                else if (arg.StartsWith("/I", StringComparison.OrdinalIgnoreCase) ||
                    arg.StartsWith("-I", StringComparison.OrdinalIgnoreCase))
                {
                    var path = arg.Substring(2);
                    if (string.IsNullOrEmpty(path) && i + 1 < args.Length)
                        path = args[++i];
                    if (!string.IsNullOrEmpty(path))
                        includes.Add(path);
                }
                else if (arg.StartsWith("/D", StringComparison.OrdinalIgnoreCase) ||
                         arg.StartsWith("-D", StringComparison.OrdinalIgnoreCase))
                {
                    var def = arg.Substring(2);
                    if (string.IsNullOrEmpty(def) && i + 1 < args.Length)
                        def = args[++i];
                    if (!string.IsNullOrEmpty(def))
                        defines.Add(def);
                }
                else if (arg.StartsWith("/std:", StringComparison.OrdinalIgnoreCase) ||
                         arg.StartsWith("-std:", StringComparison.OrdinalIgnoreCase))
                {
                    langStd = arg.Substring(5);
                }
                else if (IsRuntimeLibFlag(arg))
                {
                    runtimeLib = arg.TrimStart('/').TrimStart('-');
                }
                else if (IsWarningFlag(arg) || IsOptimizationFlag(arg))
                {
                    // Drop warning and optimization flags
                }
                else if (!arg.StartsWith("/") && !arg.StartsWith("-") &&
                         (arg.EndsWith(".cpp", StringComparison.OrdinalIgnoreCase) ||
                          arg.EndsWith(".c", StringComparison.OrdinalIgnoreCase) ||
                          arg.EndsWith(".cc", StringComparison.OrdinalIgnoreCase) ||
                          arg.EndsWith(".cxx", StringComparison.OrdinalIgnoreCase)))
                {
                    sourceFile = arg;
                }
                else
                {
                    other.Add(arg);
                }
            }

            return new ParsedArgs(clExe, sourceFile, includes, defines, langStd, runtimeLib, other);
        }

        private static bool IsWarningFlag(string arg)
        {
            var a = arg.TrimStart('/').TrimStart('-');
            // /W0-4, /Wall, /WX, /wd*, /we*, /wo*, /external:*
            if (a.StartsWith("W", StringComparison.OrdinalIgnoreCase) && a.Length >= 2)
                return true;
            if (a.StartsWith("wd", StringComparison.OrdinalIgnoreCase) ||
                a.StartsWith("we", StringComparison.OrdinalIgnoreCase) ||
                a.StartsWith("wo", StringComparison.OrdinalIgnoreCase))
                return true;
            if (a.StartsWith("external:W", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private static bool IsOptimizationFlag(string arg)
        {
            var a = arg.TrimStart('/').TrimStart('-');
            // /O1, /O2, /Od, /Ox, /Ob*, /GL, /Gw, /Oi, /Ot, /Os, /Oy
            if (a.Equals("GL", StringComparison.OrdinalIgnoreCase) ||
                a.Equals("Gw", StringComparison.OrdinalIgnoreCase))
                return true;
            if (a.StartsWith("O", StringComparison.OrdinalIgnoreCase) && a.Length >= 2)
            {
                char second = a[1];
                // /O1, /O2, /Od, /Ox, /Ob0-3, /Oi, /Ot, /Os, /Oy
                if (char.IsLetterOrDigit(second))
                    return true;
            }
            return false;
        }

        private static bool IsRuntimeLibFlag(string arg)
        {
            var a = arg.TrimStart('/').TrimStart('-');
            return a.Equals("MT", StringComparison.OrdinalIgnoreCase) ||
                   a.Equals("MTd", StringComparison.OrdinalIgnoreCase) ||
                   a.Equals("MD", StringComparison.OrdinalIgnoreCase) ||
                   a.Equals("MDd", StringComparison.OrdinalIgnoreCase);
        }

        #endregion

        #region Merge Strategies

        private static List<string> MergeDefines(List<List<string>> allDefines)
        {
            // Parse each define into name=value pairs
            var allParsed = allDefines
                .Select(defs => defs.Select(ParseDefine).ToList())
                .ToList();

            // Collect all define names
            var allNames = new LinkedHashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var parsed in allParsed)
                foreach (var (name, _) in parsed)
                    allNames.Add(name);

            var result = new List<string>();
            bool hasDebug = false;

            foreach (var name in allNames)
            {
                // Skip _DEBUG and NDEBUG, we'll add _DEBUG at the end
                if (name.Equals("_DEBUG", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("NDEBUG", StringComparison.OrdinalIgnoreCase))
                {
                    hasDebug = true;
                    continue;
                }

                // Collect all values for this define across entries
                var values = new List<string?>();
                foreach (var parsed in allParsed)
                {
                    var match = parsed.FirstOrDefault(d =>
                        d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    if (match.Name != null)
                        values.Add(match.Value);
                }

                if (values.Count == 0) continue;

                // All same? Keep that value. Conflict? Keep the define (permissive).
                var distinct = values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (distinct.Count == 1)
                {
                    result.Add(FormatDefine(name, distinct[0]));
                }
                else
                {
                    // Conflict: keep define (first non-null value, permissive)
                    var firstValue = values.FirstOrDefault(v => v != null) ?? values[0];
                    result.Add(FormatDefine(name, firstValue));
                }
            }

            // Add _DEBUG or NDEBUG based on the preferred (first) entry's defines
            if (hasDebug || allParsed.Any(p => p.Any(d =>
                d.Name.Equals("_DEBUG", StringComparison.OrdinalIgnoreCase) ||
                d.Name.Equals("NDEBUG", StringComparison.OrdinalIgnoreCase))))
            {
                var preferredDefines = allParsed[0];
                bool preferredHasDebug = preferredDefines.Any(d =>
                    d.Name.Equals("_DEBUG", StringComparison.OrdinalIgnoreCase));

                if (preferredHasDebug)
                    result.Add("_DEBUG");
                else
                    result.Add("NDEBUG");
            }

            return result;
        }

        private static (string Name, string? Value) ParseDefine(string define)
        {
            var eq = define.IndexOf('=');
            if (eq < 0) return (define, null);
            return (define.Substring(0, eq), define.Substring(eq + 1));
        }

        private static string FormatDefine(string name, string? value)
        {
            return value != null ? $"{name}={value}" : name;
        }

        private static string? PickHighestStandard(IEnumerable<string> standards)
        {
            string? best = null;
            int bestRank = -1;

            foreach (var std in standards)
            {
                int rank = GetStandardRank(std);
                if (rank > bestRank)
                {
                    bestRank = rank;
                    best = std;
                }
            }

            return best;
        }

        private static int GetStandardRank(string std)
        {
            for (int i = 0; i < CppStandardRank.Length; i++)
                if (CppStandardRank[i].Equals(std, StringComparison.OrdinalIgnoreCase))
                    return i + 100; // C++ standards ranked higher than C

            for (int i = 0; i < CStandardRank.Length; i++)
                if (CStandardRank[i].Equals(std, StringComparison.OrdinalIgnoreCase))
                    return i;

            return -1; // unknown standard
        }

        private static string? PickDebugRuntime(IEnumerable<string> runtimes)
        {
            var all = runtimes.ToList();
            if (all.Count == 0) return null;

            // Prefer debug variants
            if (all.Any(r => r.Equals("MDd", StringComparison.OrdinalIgnoreCase)))
                return "MDd";
            if (all.Any(r => r.Equals("MTd", StringComparison.OrdinalIgnoreCase)))
                return "MTd";
            if (all.Any(r => r.Equals("MD", StringComparison.OrdinalIgnoreCase)))
                return "MD";
            if (all.Any(r => r.Equals("MT", StringComparison.OrdinalIgnoreCase)))
                return "MT";

            return all[0];
        }

        private static List<string> UnionOrderPreserving(
            IEnumerable<string> items,
            StringComparer comparer)
        {
            var seen = new HashSet<string>(comparer);
            var result = new List<string>();
            foreach (var item in items)
            {
                if (seen.Add(item))
                    result.Add(item);
            }
            return result;
        }

        private static List<string> StripWarningFlags(List<string> args)
        {
            return args.Where(a => !IsWarningFlag(a)).ToList();
        }

        private static List<string> StripOptimizationFlags(List<string> args)
        {
            return args.Where(a => !IsOptimizationFlag(a)).ToList();
        }

        #endregion

        /// <summary>
        /// A simple order-preserving set backed by a HashSet + List.
        /// </summary>
        private class LinkedHashSet<T> : IEnumerable<T>
        {
            private readonly HashSet<T> _set;
            private readonly List<T> _list = new();

            public LinkedHashSet(IEqualityComparer<T>? comparer = null)
            {
                _set = new HashSet<T>(comparer);
            }

            public bool Add(T item)
            {
                if (_set.Add(item))
                {
                    _list.Add(item);
                    return true;
                }
                return false;
            }

            public IEnumerator<T> GetEnumerator() => _list.GetEnumerator();
            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
