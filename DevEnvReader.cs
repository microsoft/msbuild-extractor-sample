namespace MSBuild.CompileCommands.Extractor
{
    public class DevEnvReader
    {
        public string? VCToolsInstallDir { get; private set; }
        public string? VCTargetsPath { get; private set; }
        public string? VSInstallDir { get; private set; }
        public string? VCInstallDir { get; private set; }
        public string? WindowsSdkDir { get; private set; }
        public string? WindowsSdkVersion { get; private set; }
        public string? UCRTVersion { get; private set; }
        public string? VCToolsVersion { get; private set; }
        public string? VisualStudioVersion { get; private set; }
        public string? IncludePath { get; private set; }
        public string? ExternalIncludePath { get; private set; }
        public string? LibPath { get; private set; }
        public bool IsDevEnvAvailable { get; private set; }

        public DevEnvReader()
        {
            ReadEnvironment();
        }

        private static string? TrimTrailingBackslash(string? path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            return path.TrimEnd('\\');
        }

        private void ReadEnvironment()
        {
            VCToolsInstallDir = TrimTrailingBackslash(Environment.GetEnvironmentVariable("VCToolsInstallDir"));
            VSInstallDir = TrimTrailingBackslash(Environment.GetEnvironmentVariable("VSINSTALLDIR"));
            VCInstallDir = TrimTrailingBackslash(Environment.GetEnvironmentVariable("VCINSTALLDIR"));
            WindowsSdkDir = TrimTrailingBackslash(Environment.GetEnvironmentVariable("WindowsSdkDir"));
            WindowsSdkVersion = TrimTrailingBackslash(Environment.GetEnvironmentVariable("WindowsSDKVersion"));
            UCRTVersion = Environment.GetEnvironmentVariable("UCRTVersion");
            VCToolsVersion = Environment.GetEnvironmentVariable("VCToolsVersion");
            VisualStudioVersion = Environment.GetEnvironmentVariable("VisualStudioVersion");
            IncludePath = Environment.GetEnvironmentVariable("INCLUDE");
            ExternalIncludePath = Environment.GetEnvironmentVariable("EXTERNAL_INCLUDE");
            LibPath = Environment.GetEnvironmentVariable("LIB");

            VCTargetsPath = TrimTrailingBackslash(Environment.GetEnvironmentVariable("VCTargetsPath"));
            if (string.IsNullOrEmpty(VCTargetsPath) && !string.IsNullOrEmpty(VSInstallDir))
            {
                var possiblePaths = new[]
                {
                    Path.Combine(VSInstallDir, "MSBuild", "Microsoft", "VC", "v180"),
                    Path.Combine(VSInstallDir, "MSBuild", "Microsoft", "VC", "v170"),
                    Path.Combine(VSInstallDir, "MSBuild", "Microsoft", "VC", "v160"),
                    Path.Combine(VSInstallDir, "MSBuild", "Microsoft", "VC", "v150"),
                    Path.Combine(VSInstallDir, "Common7", "IDE", "VC", "VCTargets"),
                };

                foreach (var path in possiblePaths)
                {
                    if (Directory.Exists(path))
                    {
                        VCTargetsPath = path;
                        break;
                    }
                }
            }

            IsDevEnvAvailable = !string.IsNullOrEmpty(VCToolsInstallDir) ||
                                !string.IsNullOrEmpty(VCTargetsPath) ||
                                !string.IsNullOrEmpty(VSInstallDir);
        }

        public void ApplyToEnvironment()
        {
            if (!string.IsNullOrEmpty(VCTargetsPath))
            {
                var path = VCTargetsPath.EndsWith('\\') ? VCTargetsPath : VCTargetsPath + "\\";
                Environment.SetEnvironmentVariable("VCTargetsPath", path);
            }
        }

        public void PrintDiagnostics()
        {
            Console.WriteLine("Developer Environment Detection:");
            Console.WriteLine($"  VisualStudioVersion: {VisualStudioVersion ?? "(not set)"}");
            Console.WriteLine($"  VSINSTALLDIR:        {VSInstallDir ?? "(not set)"}");
            Console.WriteLine($"  VCINSTALLDIR:        {VCInstallDir ?? "(not set)"}");
            Console.WriteLine($"  VCToolsInstallDir:   {VCToolsInstallDir ?? "(not set)"}");
            Console.WriteLine($"  VCToolsVersion:      {VCToolsVersion ?? "(not set)"}");
            Console.WriteLine($"  VCTargetsPath:       {VCTargetsPath ?? "(not set)"}{(Environment.GetEnvironmentVariable("VCTargetsPath") == null && VCTargetsPath != null ? " (derived)" : "")}");
            Console.WriteLine($"  WindowsSdkDir:       {WindowsSdkDir ?? "(not set)"}");
            Console.WriteLine($"  WindowsSDKVersion:   {WindowsSdkVersion ?? "(not set)"}");
            Console.WriteLine($"  UCRTVersion:         {UCRTVersion ?? "(not set)"}");
            Console.WriteLine($"  INCLUDE:             {(string.IsNullOrEmpty(IncludePath) ? "(not set)" : $"({IncludePath.Split(';').Length} paths)")}");
            Console.WriteLine($"  EXTERNAL_INCLUDE:    {(string.IsNullOrEmpty(ExternalIncludePath) ? "(not set)" : $"({ExternalIncludePath.Split(';').Length} paths)")}");
            Console.WriteLine($"  LIB:                 {(string.IsNullOrEmpty(LibPath) ? "(not set)" : $"({LibPath.Split(';').Length} paths)")}");
            Console.WriteLine($"  DevEnv Available:    {IsDevEnvAvailable}");
            Console.WriteLine();
        }
    }
}
