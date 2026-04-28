using System.Diagnostics;

namespace MSBuild.CompileCommands.Extractor
{
    // Runs cl.exe against every compile command produced by the extractor and
    // reports pass/fail counts. Used by --validate to prove that extracted
    // command lines actually compile.
    internal static class ClExeValidator
    {
        public static void Run(List<CompileCommand> commands)
        {
            Console.WriteLine();
            Console.WriteLine($"Validating {commands.Count} compile commands with cl.exe...");

            int passed = 0, failed = 0, skipped = 0;
            var failures = new List<(string File, int ExitCode, string Error)>();

            foreach (var cmd in commands)
            {
                var fileName = Path.GetFileName(cmd.File);
                var ext = Path.GetExtension(cmd.File).ToLowerInvariant();

                // Skip C++20 module files, they need /interface flag
                if (ext is ".ixx" or ".cppm")
                {
                    skipped++;
                    continue;
                }

                Console.Write($"\r  [{passed + failed + skipped}/{commands.Count}] {fileName,-50}");

                var args = PrepareArgs(cmd.Arguments);

                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = args[0],
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WorkingDirectory = cmd.Directory
                    };

                    foreach (var arg in args.Skip(1))
                        startInfo.ArgumentList.Add(arg);

                    using var process = Process.Start(startInfo);
                    if (process == null)
                    {
                        failed++;
                        failures.Add((fileName, -1, "Failed to start cl.exe"));
                        continue;
                    }

                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();

                    if (!process.WaitForExit(60000))
                    {
                        try { process.Kill(true); } catch { }
                        failed++;
                        failures.Add((fileName, -1, "TIMEOUT (60s)"));
                        continue;
                    }

                    var stdout = stdoutTask.Result;
                    var stderr = stderrTask.Result;

                    if (process.ExitCode == 0)
                    {
                        passed++;
                    }
                    else
                    {
                        failed++;
                        // cl.exe writes compilation errors to stdout, banner to stderr
                        var allOutput = stdout + "\n" + stderr;
                        var errorLine = (allOutput.Split('\n')
                            .FirstOrDefault(l => l.Contains("error C", StringComparison.OrdinalIgnoreCase)
                                              || l.Contains("fatal error", StringComparison.OrdinalIgnoreCase))
                            ?? allOutput.Split('\n').FirstOrDefault(l => l.Contains("error", StringComparison.OrdinalIgnoreCase))
                            ?? "unknown error").Trim();
                        failures.Add((fileName, process.ExitCode, errorLine[..Math.Min(120, errorLine.Length)]));
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    failures.Add((fileName, -1, ex.Message[..Math.Min(120, ex.Message.Length)]));

                    if (passed + failed == 1 && ex is System.ComponentModel.Win32Exception)
                    {
                        Console.WriteLine();
                        Console.Error.WriteLine("Error: cl.exe not found. Run from a Developer Command Prompt or ensure cl.exe is in PATH.");
                        return;
                    }
                }
            }

            PrintSummary(commands.Count, passed, failed, skipped, failures);
        }

        // Transforms an extracted argument list into one safe to pass to cl.exe
        // for a smoke compile. Strips PDB/PCH side effects, normalizes DEBUG
        // macros, removes clang-only flags and warning controls.
        private static List<string> PrepareArgs(IReadOnlyList<string> source)
        {
            var args = new List<string>(source);

            // Ensure /c is present
            if (!args.Any(a => a.Equals("/c", StringComparison.OrdinalIgnoreCase)))
                args.Insert(1, "/c");

            // Replace /Fo with /FoNUL to discard .obj output
            for (int i = 0; i < args.Count; i++)
            {
                if (args[i].StartsWith("/Fo", StringComparison.OrdinalIgnoreCase) &&
                    !args[i].StartsWith("/Fp", StringComparison.OrdinalIgnoreCase))
                    args[i] = "/FoNUL";
            }

            // Remove /Fd (PDB) to avoid file locking
            args.RemoveAll(a => a.StartsWith("/Fd", StringComparison.OrdinalIgnoreCase));

            // Strip precompiled header flags. PCH compatibility is too fragile
            // with the extractor's additional flags (system /I paths, clang flags).
            args.RemoveAll(a => a.StartsWith("/Yu", StringComparison.OrdinalIgnoreCase) ||
                                a.StartsWith("/Yc", StringComparison.OrdinalIgnoreCase) ||
                                (a.StartsWith("/Fp", StringComparison.OrdinalIgnoreCase) &&
                                !a.StartsWith("/fp:", StringComparison.OrdinalIgnoreCase)));

            // Fix DEBUG macro: some headers define DEBUG without a value
            // (`#define DEBUG`) when _DEBUG is set, while others use `#if DEBUG`
            // which requires a numeric constant. Without a PCH mediating include
            // order, this causes C1017. Ensure DEBUG=1 when _DEBUG is present so
            // both patterns work.
            bool hasUnderscoreDebug = args.Any(a => a.Equals("/D_DEBUG", StringComparison.OrdinalIgnoreCase) ||
                a.Equals("_DEBUG", StringComparison.OrdinalIgnoreCase));
            if (hasUnderscoreDebug)
            {
                for (int i = args.Count - 1; i >= 0; i--)
                {
                    if (i + 1 < args.Count &&
                        args[i].Equals("/D", StringComparison.OrdinalIgnoreCase) &&
                        args[i + 1].Equals("DEBUG", StringComparison.OrdinalIgnoreCase))
                    {
                        args[i + 1] = "DEBUG=1";
                    }
                    else if (args[i].Equals("/DDEBUG", StringComparison.OrdinalIgnoreCase))
                    {
                        args[i] = "/DDEBUG=1";
                    }
                }
                if (!args.Any(a => a.Contains("DEBUG=1", StringComparison.OrdinalIgnoreCase) ||
                                   a.Contains("DEBUG=", StringComparison.OrdinalIgnoreCase)))
                    args.Add("/DDEBUG=1");
            }

            // Strip clang-only flags that the extractor injects for clangd.
            // cl.exe emits D9002 on each unknown option, which becomes a hard
            // error under /WX. These flags carry no semantic value for cl.exe.
            args.RemoveAll(a => a.StartsWith("-ferror-limit", StringComparison.Ordinal) ||
                                a.StartsWith("--target=", StringComparison.Ordinal) ||
                                a.StartsWith("-fms-compatibility", StringComparison.Ordinal) ||
                                a.StartsWith("-fdelayed-template", StringComparison.Ordinal) ||
                                a.StartsWith("-Wno-", StringComparison.Ordinal) ||
                                (a.StartsWith("-W", StringComparison.Ordinal) && a.Length > 2));

            // Convert /external:I to /I for validator compatibility.
            // cl.exe accepts /external:I natively only with /experimental:external,
            // and the two-arg form (`/external:I <path>`) confuses argument tokenization
            // on CMake-generated command lines where the path is a separate token.
            for (int i = 0; i < args.Count; i++)
            {
                if (args[i].Equals("/external:I", StringComparison.OrdinalIgnoreCase) &&
                    i + 1 < args.Count)
                {
                    args[i] = "/I" + args[i + 1];
                    args.RemoveAt(i + 1);
                }
                else if (args[i].StartsWith("/external:I", StringComparison.OrdinalIgnoreCase) &&
                         args[i].Length > "/external:I".Length)
                {
                    args[i] = "/I" + args[i]["/external:I".Length..];
                }
            }

            // Suppress all warnings during validation: we only care whether the
            // command can be parsed and the file compiled, not warning hygiene.
            // Drop existing warning-control flags, then add /w to silence everything.
            args.RemoveAll(a =>
                (a.StartsWith("/W", StringComparison.Ordinal) && a.Length > 2 && a != "/WX-") ||
                a.Equals("/WX", StringComparison.Ordinal) ||
                a.StartsWith("/wd", StringComparison.Ordinal) ||
                a.StartsWith("/we", StringComparison.Ordinal) ||
                a.StartsWith("/external:W", StringComparison.OrdinalIgnoreCase) ||
                a.Equals("/experimental:external", StringComparison.OrdinalIgnoreCase) ||
                a.StartsWith("/analyze", StringComparison.OrdinalIgnoreCase));
            args.Add("/w");

            return args;
        }

        private static void PrintSummary(int total, int passed, int failed, int skipped,
            List<(string File, int ExitCode, string Error)> failures)
        {
            Console.Write($"\r{new string(' ', 70)}\r");
            Console.WriteLine();
            Console.WriteLine(new string('=', 60));
            Console.WriteLine($"VALIDATION: {passed} passed, {failed} failed, {skipped} skipped out of {total}");
            Console.WriteLine(new string('=', 60));

            if (skipped > 0)
                Console.WriteLine($"  Skipped {skipped} module file(s) (.ixx/.cppm)");

            if (failures.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("Failed:");
                foreach (var (file, exitCode, error) in failures.Take(20))
                    Console.WriteLine($"  [{exitCode}] {file}: {error}");
                if (failures.Count > 20)
                    Console.WriteLine($"  ... and {failures.Count - 20} more");
            }
            else if (failed == 0)
            {
                Console.WriteLine();
                Console.WriteLine("All files compiled successfully!");
            }
        }
    }
}
