using System.Text;

namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Tokenizes command line strings using Windows CommandLineToArgvW semantics.
    /// Handles quoted strings, escaped quotes, backslash rules, and @response files.
    /// </summary>
    public static class CommandLineTokenizer
    {
        /// <summary>
        /// Tokenizes a command line string using Windows CommandLineToArgvW semantics.
        /// </summary>
        public static List<string> Tokenize(string commandLine)
        {
            var tokens = new List<string>();
            if (string.IsNullOrWhiteSpace(commandLine))
                return tokens;

            int i = 0;
            while (i < commandLine.Length)
            {
                while (i < commandLine.Length && char.IsWhiteSpace(commandLine[i]))
                    i++;

                if (i >= commandLine.Length)
                    break;

                tokens.Add(ParseToken(commandLine, ref i));
            }

            return tokens;
        }

        private static string ParseToken(string commandLine, ref int i)
        {
            var token = new StringBuilder();
            bool inQuotes = false;

            while (i < commandLine.Length && (inQuotes || !char.IsWhiteSpace(commandLine[i])))
            {
                int numBackslashes = 0;
                while (i < commandLine.Length && commandLine[i] == '\\')
                {
                    i++;
                    numBackslashes++;
                }

                if (i < commandLine.Length && commandLine[i] == '"')
                {
                    // Backslashes followed by a double quote:
                    // - Even count: half become literal backslashes, quote toggles mode
                    // - Odd count: half become literal backslashes, last one escapes the quote
                    token.Append('\\', numBackslashes / 2);
                    if (numBackslashes % 2 == 0)
                    {
                        // "" inside a quoted region produces a literal " (CommandLineToArgvW behavior)
                        if (inQuotes && i + 1 < commandLine.Length && commandLine[i + 1] == '"')
                        {
                            token.Append('"');
                            i += 2;
                        }
                        else
                        {
                            inQuotes = !inQuotes;
                            i++;
                        }
                    }
                    else
                    {
                        token.Append('"');
                        i++;
                    }
                }
                else
                {
                    // Backslashes not followed by a quote are literal
                    token.Append('\\', numBackslashes);
                    if (i < commandLine.Length && (inQuotes || !char.IsWhiteSpace(commandLine[i])))
                    {
                        token.Append(commandLine[i]);
                        i++;
                    }
                }
            }

            return token.ToString();
        }

        /// <summary>
        /// Tokenizes a command line and expands @response_file references by
        /// reading and recursively tokenizing their contents.
        /// </summary>
        public static List<string> TokenizeWithResponseFiles(string commandLine, string? workingDirectory = null)
        {
            return TokenizeWithResponseFilesCore(commandLine, workingDirectory, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0);
        }

        private const int MaxResponseFileDepth = 10;

        private static List<string> TokenizeWithResponseFilesCore(string commandLine, string? workingDirectory, HashSet<string> visitedFiles, int depth)
        {
            var tokens = Tokenize(commandLine);
            var expanded = new List<string>();

            foreach (var token in tokens)
            {
                if (token.StartsWith('@') && token.Length > 1)
                {
                    var filePath = token[1..];
                    if (!Path.IsPathRooted(filePath) && workingDirectory != null)
                        filePath = Path.Combine(workingDirectory, filePath);

                    string normalizedPath;
                    try
                    {
                        normalizedPath = Path.GetFullPath(filePath);
                    }
                    catch (Exception)
                    {
                        // Treat as literal token on malformed path
                        expanded.Add(token);
                        continue;
                    }

                    if (depth >= MaxResponseFileDepth || visitedFiles.Contains(normalizedPath))
                    {
                        // Cycle detected or max depth exceeded; treat as literal token
                        expanded.Add(token);
                        continue;
                    }

                    if (File.Exists(normalizedPath))
                    {
                        try
                        {
                            visitedFiles.Add(normalizedPath);
                            var content = File.ReadAllText(normalizedPath);
                            expanded.AddRange(TokenizeWithResponseFilesCore(content, Path.GetDirectoryName(normalizedPath), visitedFiles, depth + 1));
                            visitedFiles.Remove(normalizedPath);
                        }
                        catch (Exception)
                        {
                            // Treat as literal token on any failure (malformed path, access denied, etc.)
                            expanded.Add(token);
                        }
                    }
                    else
                    {
                        expanded.Add(token);
                    }
                }
                else
                {
                    expanded.Add(token);
                }
            }

            return expanded;
        }
    }
}
