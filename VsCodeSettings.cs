using System.Text.Json;

namespace MSBuild.CompileCommands.Extractor
{
    /// <summary>
    /// Helpers that wire the extractor's output into a Visual Studio Code workspace
    /// for use with the Microsoft C/C++ extension (cpptools). These helpers are
    /// independent of the extractor output format; they only run when the user
    /// passes --c-cpp-properties (which is mutually exclusive with --format rich).
    /// </summary>
    public static class VsCodeSettings
    {
        /// <summary>
        /// Generate a .vscode/c_cpp_properties.json that references the compile_commands.json,
        /// and a .vscode/settings.json that turns on cpptools squiggles and verbose logging.
        /// </summary>
        public static void GenerateCCppProperties(string compileCommandsPath, string platform, string baseDir)
        {
            var vsCodeDir = Path.Combine(baseDir, ".vscode");
            Directory.CreateDirectory(vsCodeDir);

            var propsPath = Path.Combine(vsCodeDir, "c_cpp_properties.json");

            if (File.Exists(propsPath))
            {
                Console.WriteLine($"Warning: {propsPath} already exists, overwriting");
            }

            // Make compileCommands path relative using ${workspaceFolder} if possible
            string compileCommandsRef;
            var fullCcPath = Path.GetFullPath(compileCommandsPath);
            var fullBaseDir = Path.GetFullPath(baseDir).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (fullCcPath.StartsWith(fullBaseDir, StringComparison.OrdinalIgnoreCase))
                compileCommandsRef = "${workspaceFolder}/" + fullCcPath[fullBaseDir.Length..].Replace('\\', '/');
            else
                compileCommandsRef = fullCcPath.Replace('\\', '/');

            var intelliSenseMode = platform.ToLowerInvariant() switch
            {
                "x64" => "msvc-x64",
                "win32" or "x86" => "msvc-x86",
                "arm64" or "arm64ec" => "msvc-arm64",
                "arm" => "msvc-arm",
                _ => "msvc-x64"
            };

            var configName = platform.ToLowerInvariant() switch
            {
                "x64" => "MSVC x64",
                "win32" or "x86" => "MSVC x86",
                "arm64" => "MSVC ARM64",
                "arm" => "MSVC ARM",
                _ => "MSVC"
            };

            var config = new
            {
                configurations = new[]
                {
                    new
                    {
                        name = configName,
                        compileCommands = compileCommandsRef,
                        intelliSenseMode = intelliSenseMode
                    }
                },
                version = 4
            };

            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            File.WriteAllText(propsPath, json);
            Console.WriteLine($"Wrote {propsPath}");

            // Also generate .vscode/settings.json with C++ diagnostics enabled
            var settingsPath = Path.Combine(vsCodeDir, "settings.json");
            var settings = new Dictionary<string, object>
            {
                ["C_Cpp.errorSquiggles"] = "enabled",
                ["C_Cpp.loggingLevel"] = "Debug",
                ["C_Cpp.default.enableConfigurationSquiggles"] = true
            };

            if (File.Exists(settingsPath))
            {
                // Merge into existing settings, don't overwrite unrelated keys
                try
                {
                    var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(File.ReadAllText(settingsPath));
                    if (existing != null)
                    {
                        foreach (var kv in existing)
                        {
                            if (!settings.ContainsKey(kv.Key))
                                settings[kv.Key] = kv.Value;
                        }
                    }
                }
                catch { /* overwrite if unparseable */ }
            }

            var settingsJson = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(settingsPath, settingsJson);
            Console.WriteLine($"Wrote {settingsPath}");
        }
    }
}
