# RipDisc - C# Console Application

This is a C# port of the PowerShell `rip-disc.ps1` script. It covers the core rip/encode/organize workflow for DVD and Blu-ray discs using MakeMKV and HandBrake; newer PowerShell-only features are listed in the Feature Parity table in the main README.

## Solution Layout

| Project | Purpose |
|---------|---------|
| `RipDisc.Core` | The pipeline (rip, encode, organize, open), config loading, queue, MakeMKV error analysis. No console code: all output and prompts go through the `IRipUI` interface, so a GUI front end can drive the same pipeline. |
| `RipDisc.Cli` | The console app (`RipDisc.exe`): argument parsing and `ConsoleRipUI`, the console implementation of `IRipUI`. |
| `RipDisc.Tests` | xUnit tests, including end-to-end pipeline runs against fake MakeMKV/HandBrake in a temp folder (no drive or tools needed). |

## Requirements

- .NET 8.0 or later
- Windows OS
- MakeMKV and HandBrakeCLI installed. Paths come from `ripdisc-config.json` (the same file the PowerShell scripts use), found next to `RipDisc.exe` or in any parent folder; otherwise the usual install locations and `PATH` are searched.

## Configuration

`ripdisc-config.json` (see `ripdisc-config.sample.json` in the repo root) sets:

- `makemkvPath`, `handbrakePath` - tool locations (auto-detected when blank)
- `tempRoot` - where MakeMKV writes, logs go (`<tempRoot>\logs`) and the encode queue lives (`<tempRoot>\handbrake-queue.json`); default `C:\Video`
- `defaultInputDrive`, `defaultOutputDrive` - used when `-drive` / `-outputDrive` are not given; default `D:` / `E:`
- `driveLabels` - names shown for `-driveIndex` drives

## Building

From the `RipDisc` directory:

```bash
dotnet build RipDisc.sln -c Release
dotnet test RipDisc.sln
```

The executable will be located at:
```
RipDisc.Cli\bin\Release\net8.0-windows\RipDisc.exe
```

`nuget.config` in this folder restores from nuget.org only.

## Usage

```bash
RipDisc -title <title> [options]
```

### Required Parameters

- `-title <string>` - Title of the movie or series

### Optional Parameters

- `-series` - Flag for TV series (no value needed)
- `-season <int>` - Season number (default: 0)
- `-disc <int>` - Disc number (default: 1)
- `-drive <string>` - Drive letter (default: `defaultInputDrive` from config, else D:)
- `-driveIndex <int>` - Drive index for MakeMKV (default: -1)
- `-outputDrive <string>` - Output drive letter (default: `defaultOutputDrive` from config, else E:). `F`, `F:` and `F:\` all work.
- `-queue` - Rip now, add the encode to the shared queue
- `-processQueue` - Encode every queued job, one at a time
- `-bluray` - If an encode with subtitles fails, retry without them (Blu-ray PGS subtitles)

Prompts (start confirmation, title warnings, existing-folder choice) treat closed input as "no", so piping input never starts a rip by accident. Ctrl+C once stops MakeMKV/HandBrake and prints what is left to do; a second Ctrl+C exits immediately.

### Examples

Rip a movie:
```bash
RipDisc -title "The Matrix"
```

Rip a TV series:
```bash
RipDisc -title "Breaking Bad" -series -season 1 -disc 1
```

Rip special features (disc 2):
```bash
RipDisc -title "The Matrix" -disc 2
```

Use a specific drive index:
```bash
RipDisc -title "The Matrix" -driveIndex 1 -outputDrive F:
```

## Features

Implemented (the PowerShell script has more - see the Feature Parity table in the main README):

1. **Command-line argument parsing** with validation
2. **4-step processing workflow:**
   - Step 1: MakeMKV ripping to MKV files
   - Step 2: HandBrake encoding to MP4
   - Step 3: File organization (renaming, prefixing, extras folder management)
   - Step 4: Open output directory
3. **Step tracking** with completion summary
4. **Colored console output** matching PowerShell colors
5. **Comprehensive logging** to `<tempRoot>\logs\` (default `C:\Video\logs\`)
6. **Drive readiness checks** before operations
7. **MakeMKV error analysis** with specific error messages
8. **Interactive prompts** for confirmation and conflict resolution
9. **File conflict handling** with unique file path generation
10. **Window title management** for tracking concurrent rips
11. **Disc ejection** after successful rip
12. **Detailed error handling** with manual recovery guidance
13. **Movie vs TV Series workflows**
14. **Feature file identification** (largest file for movies)
15. **Extras folder organization** for non-feature content
16. **Special Features naming** for disc 2+ files

## Directory Structure

The application creates and manages the following directory structure:

**Movies:**
```
E:\DVDs\MovieName\
├── MovieName-Feature.mp4
└── extras\
    ├── MovieName-trailer.mp4
    └── MovieName-deleted-scenes.mp4
```

**TV Series (with season):**
```
E:\Series\SeriesName\
└── Season 1\
    ├── SeriesName-episode1.mp4
    └── SeriesName-episode2.mp4
```

**TV Series (no season):**
```
E:\Series\SeriesName\
├── SeriesName-episode1.mp4
└── SeriesName-episode2.mp4
```

## Temporary Files

MakeMKV temporary files are stored under `tempRoot` (default `C:\Video`):
- Disc 1: `C:\Video\TitleName\`
- Disc 2+: `C:\Video\TitleName\Disc2\`, `C:\Video\TitleName\Disc3\`, etc.

These directories are automatically cleaned up after successful encoding.

## Logs

Session logs are saved to:
```
<tempRoot>\logs\{title}_disc{disc}_{timestamp}.log
```

## Error Handling

If an error occurs:
- The window title shows `-ERROR` suffix
- Completed steps are shown in green
- Remaining steps are listed with manual instructions
- The relevant directory is opened for inspection
- Log file location is displayed

## Concurrent Ripping

The application supports concurrent ripping of multiple discs:
- Each disc uses a separate temporary directory
- Window titles identify which disc is being processed
- Status suffixes indicate processing state: `-INPUT`, `-ERROR`, `-DONE`
- For movies, disc 2+ shows `-extras` in the window title

## Differences from PowerShell Script

The C# version covers the core workflow of the PowerShell script (the Feature Parity table in the main README lists what it lacks), with these implementation differences:

1. **Process execution**: Uses `System.Diagnostics.Process` instead of PowerShell cmdlets
2. **COM interop**: Uses C# dynamic types for Shell.Application (disc ejection)
3. **File operations**: Uses `System.IO` classes instead of PowerShell file cmdlets
4. **Cross-platform safety**: Uses `[SupportedOSPlatform("windows")]` attributes

## Notes

- The original PowerShell script (`rip-disc.ps1`) remains in the repository root
- Both versions can be used interchangeably
- The C# version may offer better performance and easier distribution as a standalone executable
