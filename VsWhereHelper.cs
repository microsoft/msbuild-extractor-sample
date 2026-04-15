using System.Runtime.InteropServices;
using System.Text.Json;

namespace MSBuild.CompileCommands.Extractor
{
    public class VsWhereResult
    {
        public string? InstallationPath { get; set; }
        public string? VCTargetsPath { get; set; }
        public string? VCToolsInstallDir { get; set; }
    }

    public record VsInstance(
        string InstanceId,
        string DisplayName,
        string InstallationPath,
        string Version,
        bool HasVcTools,
        string? VCTargetsPath,
        string? VCToolsInstallDir
    );

    /// <summary>
    /// Discovers Visual Studio installations and VC++ tool paths using vswhere.exe.
    /// Also resolves the real cl.exe path from VCToolsInstallDir.
    /// </summary>
    public static class VsWhereHelper
    {
        private static readonly string[] VCTargetVersionDirs = ["v180", "v170", "v160", "v150"];

        /// <summary>
        /// Lists all Visual Studio installations found by vswhere.exe.
        /// </summary>
        public static List<VsInstance> ListInstances(bool enableLogger = false)
        {
            var results = new List<VsInstance>();
            var vswhereExe = FindVsWhere();
            if (vswhereExe == null)
            {
                if (enableLogger)
                    Console.WriteLine("VsWhere: vswhere.exe not found");
                return results;
            }

            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = vswhereExe,
                    Arguments = "-all -prerelease -format json",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(startInfo);
                if (process == null) return results;

                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();

                if (!process.WaitForExit(10000))
                {
                    try { process.Kill(true); } catch { }
                    return results;
                }

                var output = stdoutTask.Result;
                // Consume stderr to prevent hang
                _ = stderrTask.Result;

                if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
                    return results;

                using var doc = JsonDocument.Parse(output);
                var root = doc.RootElement;

                if (root.ValueKind != JsonValueKind.Array)
                    return results;

                foreach (var entry in root.EnumerateArray())
                {
                    var instanceId = entry.TryGetProperty("instanceId", out var idEl) ? idEl.GetString() : null;
                    var displayName = entry.TryGetProperty("displayName", out var nameEl) ? nameEl.GetString() : null;
                    var installPath = entry.TryGetProperty("installationPath", out var pathEl) ? pathEl.GetString() : null;
                    var version = entry.TryGetProperty("installationVersion", out var verEl) ? verEl.GetString() : null;

                    if (instanceId == null || installPath == null)
                        continue;

                    var vcTargetsPath = FindVCTargetsPath(installPath);
                    var vcToolsInstallDir = FindVCToolsInstallDir(installPath);
                    var hasVcTools = vcTargetsPath != null || vcToolsInstallDir != null;

                    results.Add(new VsInstance(
                        InstanceId: instanceId,
                        DisplayName: displayName ?? "(unknown)",
                        InstallationPath: installPath,
                        Version: version ?? "(unknown)",
                        HasVcTools: hasVcTools,
                        VCTargetsPath: vcTargetsPath,
                        VCToolsInstallDir: vcToolsInstallDir
                    ));
                }
            }
            catch (Exception ex)
            {
                if (enableLogger)
                    Console.WriteLine($"VsWhere: Error listing instances: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// Auto-detects a Visual Studio installation with C++ tools via vswhere.exe.
        /// Returns null if vswhere is not found or no suitable installation exists.
        /// </summary>
        public static VsWhereResult? DetectVisualStudio(bool enableLogger = false)
        {
            var vswhereExe = FindVsWhere();
            if (vswhereExe == null)
            {
                if (enableLogger)
                    Console.WriteLine("VsWhere: vswhere.exe not found");
                return null;
            }

            var installPath = RunVsWhere(vswhereExe, enableLogger);
            if (installPath == null)
            {
                if (enableLogger)
                    Console.WriteLine("VsWhere: No Visual Studio installation with C++ tools found");
                return null;
            }

            var result = new VsWhereResult { InstallationPath = installPath };
            result.VCTargetsPath = FindVCTargetsPath(installPath);
            result.VCToolsInstallDir = FindVCToolsInstallDir(installPath);

            if (enableLogger)
            {
                Console.WriteLine("VsWhere Detection:");
                Console.WriteLine($"  InstallationPath:  {result.InstallationPath}");
                Console.WriteLine($"  VCTargetsPath:     {result.VCTargetsPath ?? "(not found)"}");
                Console.WriteLine($"  VCToolsInstallDir: {result.VCToolsInstallDir ?? "(not found)"}");
                Console.WriteLine();
            }

            return result;
        }

        private static string? FindVsWhere()
        {
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (!string.IsNullOrEmpty(programFilesX86))
            {
                var path = Path.Combine(programFilesX86, "Microsoft Visual Studio", "Installer", "vswhere.exe");
                if (File.Exists(path))
                    return path;
            }

            // Fallback: check PATH
            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "vswhere.exe",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(startInfo);
                if (process != null)
                {
                    var output = process.StandardOutput.ReadLine();
                    process.WaitForExit(5000);
                    if (process.ExitCode == 0 && !string.IsNullOrEmpty(output) && File.Exists(output))
                        return output;
                }
            }
            catch { }

            return null;
        }

        private static string? RunVsWhere(string vswhereExe, bool enableLogger)
        {
            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = vswhereExe,
                    Arguments = "-products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -latest -prerelease -format json",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(startInfo);
                if (process == null) return null;

                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();

                if (!process.WaitForExit(10000))
                {
                    try { process.Kill(true); } catch { }
                    return null;
                }

                var output = stdoutTask.Result;

                if (enableLogger)
                    Console.WriteLine($"VsWhere: Exit code {process.ExitCode}");

                if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
                    return null;

                using var doc = JsonDocument.Parse(output);
                var root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0)
                {
                    if (root[0].TryGetProperty("installationPath", out var pathElement))
                        return pathElement.GetString();
                }
            }
            catch (Exception ex)
            {
                if (enableLogger)
                    Console.WriteLine($"VsWhere: Error running vswhere: {ex.Message}");
            }

            return null;
        }

        public static string? FindVCTargetsPath(string vsInstallPath)
        {
            foreach (var version in VCTargetVersionDirs)
            {
                var path = Path.Combine(vsInstallPath, "MSBuild", "Microsoft", "VC", version);
                if (Directory.Exists(path))
                    return path;
            }

            var legacyPath = Path.Combine(vsInstallPath, "Common7", "IDE", "VC", "VCTargets");
            if (Directory.Exists(legacyPath))
                return legacyPath;

            return null;
        }

        public static string? FindVCToolsInstallDir(string vsInstallPath)
        {
            var msvcRoot = Path.Combine(vsInstallPath, "VC", "Tools", "MSVC");
            if (!Directory.Exists(msvcRoot))
                return null;

            var versions = Directory.GetDirectories(msvcRoot)
                .Select(d => new { Path = d, Name = Path.GetFileName(d) })
                .Where(d => Version.TryParse(d.Name, out _))
                .OrderByDescending(d => Version.Parse(d.Name))
                .ToList();

            return versions.Count > 0 ? versions[0].Path : null;
        }

        /// <summary>
        /// Resolves the real cl.exe path from VCToolsInstallDir and target platform.
        /// Maps the platform name to a target architecture and probes for cl.exe
        /// under the appropriate Host&lt;arch&gt;/&lt;target&gt; directory.
        /// </summary>
        public static string? ResolveClExePath(string? vcToolsInstallDir, string platform)
        {
            if (string.IsNullOrEmpty(vcToolsInstallDir))
                return null;

            var hostArch = RuntimeInformation.OSArchitecture switch
            {
                Architecture.X64 => "Hostx64",
                Architecture.X86 => "Hostx86",
                Architecture.Arm64 => "HostARM64",
                _ => "Hostx64"
            };

            var targetArch = MapPlatformToTarget(platform);

            // Try the natural host/target combination first
            var clPath = Path.Combine(vcToolsInstallDir, "bin", hostArch, targetArch, "cl.exe");
            if (File.Exists(clPath))
                return clPath;

            // Fallback: try other host architectures
            foreach (var host in new[] { "Hostx64", "Hostx86", "HostARM64" })
            {
                if (host == hostArch) continue;
                clPath = Path.Combine(vcToolsInstallDir, "bin", host, targetArch, "cl.exe");
                if (File.Exists(clPath))
                    return clPath;
            }

            return null;
        }

        private static string MapPlatformToTarget(string platform)
        {
            return platform.ToLowerInvariant() switch
            {
                "x64" => "x64",
                "win32" => "x86",
                "x86" => "x86",
                "arm" => "arm",
                "arm64" => "arm64",
                "arm64ec" => "arm64",
                _ => "x64"
            };
        }

        /// <summary>
        /// Finds the latest installed Windows 10/11 SDK version by scanning
        /// the Include directory under "Windows Kits\10".
        /// Returns null if no SDK is installed.
        /// </summary>
        public static string? FindLatestWindowsSdkVersion()
        {
            var sdkInclude = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                "Windows Kits", "10", "Include");

            if (!Directory.Exists(sdkInclude))
                return null;

            return Directory.GetDirectories(sdkInclude)
                .Select(d => Path.GetFileName(d))
                .Where(v => v != null && v.StartsWith("10.0.") && Version.TryParse(v, out _))
                .OrderByDescending(v => Version.Parse(v!))
                .FirstOrDefault();
        }
    }
}
