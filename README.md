# MSBuild Compile Commands Extractor

Extract [`compile_commands.json`](https://clang.llvm.org/docs/JSONCompilationDatabase.html) from Visual C++ MSBuild projects - the same compilation database used by [clangd](https://clangd.llvm.org/), [clang-tidy](https://clang.llvm.org/extra/clang-tidy/), and VS Code C++ extensions for IntelliSense, code navigation, and static analysis.

Works with `.vcxproj`, `.sln`, `.slnx`, and CMake-generated Visual Studio solutions.

## Examples

### Basic extraction

```bash
msbuild-extractor --project myapp.vcxproj -c Debug -a x64
```

### Solution with multiple projects

```bash
msbuild-extractor --solution myapp.sln -c Debug -a x64 -o compile_commands.json
```

### Auto-detect Visual Studio (no manual paths needed)

```bash
# List installed VS instances
msbuild-extractor --list-instances

# Use a specific VS instance
msbuild-extractor --vs-instance cdf4 --solution myapp.sln -c Debug -a x64
```

### Validate extracted commands (compile every file with cl.exe)

```bash
msbuild-extractor --solution myapp.sln -c Debug -a x64 --validate
```

### Multi-configuration extraction

```bash
# Extract all configurations (Debug/Release × x64/Win32/ARM64)
msbuild-extractor --project myapp.vcxproj --all-configurations

# Merge into single file with one entry per source file
msbuild-extractor --project myapp.vcxproj --all-configurations --merge --deduplicate
```

### Merge multiple solutions

```bash
msbuild-extractor --solution engine.sln --solution editor.sln -c Debug -a x64 -o compile_commands.json
```

### Rich hierarchical format

```bash
# Structured output with solutions/projects/configurations/files
msbuild-extractor --solution myapp.sln --format rich -o compile_database.json
```

### Out-of-process mode (custom MSBuild)

```bash
msbuild-extractor --solution myapp.sln --msbuild-path "C:\VS\MSBuild\Current\Bin\MSBuild.exe" -c Debug -a x64
```

## Features

- **Zero-config extraction** - auto-detects Visual Studio via `vswhere.exe`, resolves `VCTargetsPath`, `VCToolsInstallDir`, and real `cl.exe` path automatically
- **Two extraction modes** - in-process (MSBuild API) and out-of-process (spawns MSBuild.exe, parses binlog) for maximum compatibility
- **Multi-config** - `--all-configurations` extracts every `Configuration|Platform` combination; `--merge` combines them; `--deduplicate` produces one smart entry per file for IntelliSense
- **Multi-input** - specify multiple `--solution` and `--project` arguments to merge across solutions
- **Built-in validation** - `--validate` compiles every extracted entry with `cl.exe /c` to prove correctness
- **Config validation** - `--strict` rejects invalid configuration/platform combinations with available options listed
- **Rich format** - `--format rich` outputs a hierarchical JSON schema with solutions, projects, configurations, structured includes/defines, and toolchain metadata
- **GN project support** - extracts IntelliSense settings from GN-generated VS solutions (Chromium, Crashpad, WebRTC) via `ItemDefinitionGroup` fallback
- **VS instance selection** - `--list-instances` shows all VS installations; `--vs-instance <id>` selects by instance ID
- **clangd compatible** - clean JSON output (no extra keys), real `cl.exe` path, `-ferror-limit=0` and `--target` injection for optimal clangd experience
- **Solution formats** - `.sln`, `.slnx`, and GN-generated `.sln` files

## Installation

### Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022/2026 with C++ workload (or Build Tools)

### Build from source

```bash
git clone https://github.com/<owner>/msbuild-extractor.git
cd msbuild-extractor
dotnet build
```

### Run

```bash
dotnet run -- --solution path/to/solution.sln -c Debug -a x64
```

## Usage

```
msbuild-extractor [options]

Options:
  -p, --project <path>              Path to .vcxproj file (repeatable)
  -s, --solution <path>             Path to .sln or .slnx file (repeatable)
  -c, --configuration <name>        Build configuration [default: Debug]
  -a, --platform <name>             Build platform [default: x64]
  -o, --output <path>               Output path [default: compile_commands.json]
  --all-configurations              Extract for all configuration/platform combos
  --merge                           Merge all configs into single output file
  --deduplicate                     Smart merge: one entry per file for IntelliSense
  --prefer-configuration <name>     Preferred config for dedup conflicts [default: Debug]
  --prefer-platform <name>          Preferred platform for dedup conflicts [default: x64]
  --format <standard|rich>          Output format [default: standard]
  --validate                        Verify each entry compiles with cl.exe
  --strict                          Treat config/platform warnings as errors
  --vs-path <path>                  Path to Visual Studio installation
  --vs-instance <id>                Select VS by instance ID (see --list-instances)
  --list-instances                  List all Visual Studio installations and exit
  --vc-targets-path <path>          VC targets directory (auto-detected)
  --vc-tools-install-dir <path>     VCToolsInstallDir property (auto-detected)
  --msbuild-path <path>             MSBuild.exe path (enables out-of-process mode)
  --use-dev-env                     Read Developer Command Prompt environment variables
  --logger                          Enable MSBuild console logger output
```

At least one `--project` or `--solution` must be specified. Both can be repeated and combined.

## How It Works

The tool loads a `.vcxproj` file (or discovers all `.vcxproj` files in a solution) using the MSBuild API, runs a design-time build targeting `GetClCommandLines` (defined in `Microsoft.Cpp.DesignTime.targets`), and converts the results into the JSON Compilation Database format. No actual compilation occurs - only project evaluation and target execution.

For GN-generated projects (Chromium, Crashpad), where `GetClCommandLines` returns no items because source files aren't listed in the vcxproj, the tool falls back to reading `ItemDefinitionGroup/ClCompile` settings and scanning for source files on disk.

## Development

### Building

```bash
dotnet build
```

## Known Limitations

- **C++20 modules** (`.ixx`/`.cppm`) are extracted but skipped during `--validate` (require `/interface` flag)
- **NuGet native packages** must be restored before extraction (`nuget restore` or `msbuild -t:Restore`)
- **Build-generated files** (MIDL proxies, code generators) won't exist until a full build is done
- Windows-only - requires Visual Studio and MSVC toolchain

## License

MIT
