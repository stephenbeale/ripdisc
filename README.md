# RipDisc

PowerShell and C# tools for automated DVD and Blu-ray disc ripping using MakeMKV and HandBrake.

## Overview

This repository contains two implementations of the same disc ripping workflow:

1. **PowerShell Script** (`rip-disc.ps1`) - Original implementation
2. **C# Console Application** (`RipDisc/`) - Modern cross-language port

The PowerShell version is the primary implementation and has the most features. The C# version covers core ripping functionality but is behind on some newer features (see [Feature Parity](#feature-parity) below).

## Features

- **Automated ripping and encoding** using MakeMKV and HandBrake
- **4-step processing workflow** with progress tracking
- **Movie, TV Series, and genre-based support** (Documentary, Tutorial, Fitness, Music, Surf) with different organization strategies
- **Jellyfin episode naming** for series (`Title-S01E01.mp4`)
- **Composite mega-file detection** skips all-in-one files during series encoding
- **Multi-disc support** with concurrent ripping capability
- **HandBrake queue mode** for sequential encoding after concurrent rips
- **Blu-ray subtitle fallback** (tries subtitles first, retries without on PGS failure)
- **Feature file identification** (automatically identifies main feature)
- **Extras folder management** for special features
- **Resume failed rips** from any step with `continue-rip.ps1`
- **HandBrake recovery scripts** generated before encoding
- **Comprehensive error handling** with recovery guidance
- **Session logging** for debugging and recovery
- **Drive readiness checks** before operations
- **Interactive prompts** for confirmation and conflict resolution
- **Window title management** for tracking concurrent operations
- **Console close button protection** prevents accidental window closure
- **Automatic disc ejection** after successful rip

## Quick Start

### PowerShell Version

```powershell
.\rip-disc.ps1 -title "The Matrix"
```

### C# Version

```bash
cd RipDisc\RipDisc\bin\Release\net8.0-windows
.\RipDisc.exe -title "The Matrix"
```

## Requirements

- **Windows OS**
- **MakeMKV** installed at `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe`
- **HandBrake CLI** installed at `C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe`
- **PowerShell 5.1+** (for PowerShell version)
- **.NET 8.0+** (for C# version)

## Usage

Both versions use the same command-line parameters:

```
-title <string>         (Required) Title of the movie or series
-series                 Flag for TV series
-season <int>           Season number (default: 0)
-disc <int>             Disc number (default: 1)
-drive <string>         Drive letter (default: D:)
-driveIndex <int>       Drive index for MakeMKV (default: -1)
-outputDrive <string>   Output drive letter (default: E:)
-extras                 Flag for extras-only disc
-queue                  Queue encoding instead of running immediately
-bluray                 Blu-ray mode (subtitle fallback for PGS)
-documentary            Documentary mode (outputs to Documentaries folder)
-tutorial               Tutorial mode (outputs to Tutorials folder)
-fitness                Fitness mode (outputs to Fitness folder)
-music                  Music mode (outputs to Music folder)
-surf                   Surf mode (outputs to Surf folder)
-startEpisode <int>     Starting episode number for series (default: 1)
```

### Examples

**Rip a movie:**
```powershell
.\rip-disc.ps1 -title "The Matrix"
```

**Rip special features (disc 2):**
```powershell
.\rip-disc.ps1 -title "The Matrix" -disc 2
```

**Rip a TV series:**
```powershell
.\rip-disc.ps1 -title "Breaking Bad" -series -season 1 -disc 1
```

**Rip a TV series disc 2 (continuing episode numbers):**
```powershell
.\rip-disc.ps1 -title "Breaking Bad" -series -season 1 -disc 2 -startEpisode 5
```

**Rip a documentary:**
```powershell
.\rip-disc.ps1 -title "Planet Earth" -documentary
```

**Rip a Blu-ray:**
```powershell
.\rip-disc.ps1 -title "Inception" -bluray
```

**Queue mode for concurrent rips:**
```powershell
.\rip-disc.ps1 -title "The Matrix" -queue                        # Terminal 1
.\rip-disc.ps1 -title "The Matrix" -disc 2 -queue -driveIndex 1  # Terminal 2
RipDisc -processQueue                                             # After all rips
```

**Use specific drive index:**
```powershell
.\rip-disc.ps1 -title "The Matrix" -driveIndex 1 -outputDrive F:
```

## Directory Structure

### Movies

```
E:\DVDs\MovieName\
├── MovieName-Feature.mp4
└── extras\
    ├── MovieName-trailer.mp4
    └── MovieName-deleted-scenes.mp4
```

### TV Series (with season)

```
E:\Series\SeriesName\
└── Season 1\
    ├── SeriesName-S01E01.mp4
    ├── SeriesName-S01E02.mp4
    └── SeriesName-S01E03.mp4
```

### TV Series (no season)

```
E:\Series\SeriesName\
├── SeriesName-E01.mp4
├── SeriesName-E02.mp4
└── SeriesName-E03.mp4
```

### Documentaries

```
E:\Documentaries\DocName\
├── DocName-Feature.mp4
└── extras\
    └── DocName-bonus.mp4
```

### Music

```
E:\Music\ArtistName\
├── ArtistName-Feature.mp4
└── extras\
    └── ArtistName-behind-the-scenes.mp4
```

## Processing Steps

Both versions execute the same 4-step workflow:

1. **MakeMKV Rip** - Extract disc to MKV files
2. **HandBrake Encoding** - Encode MKV to MP4 with optimized settings
3. **Organize Files** - Rename, prefix, and organize into proper structure
4. **Open Directory** - Open output folder for verification

Each step is tracked, and the system shows completion status and provides recovery guidance if errors occur.

## Concurrent Ripping

Both versions support ripping multiple discs simultaneously:

- Each disc uses a separate temporary directory
- Window titles show which disc is being processed
- Status suffixes indicate state: `-INPUT`, `-ERROR`, `-DONE`
- For movies, disc 2+ shows `-extras` in the window title

## Logging

Session logs are saved to `C:\Video\logs\{title}_disc{disc}_{timestamp}.log`

Logs include:
- All processing steps
- File operations
- Error messages
- Recovery information

## Error Handling

If an error occurs:
- Window title shows `-ERROR` suffix
- Completed steps are displayed in green
- Remaining steps are listed with manual instructions
- Relevant directory is opened for inspection
- Log file location is provided

## Feature Parity

The PowerShell scripts are the primary implementation. The C# version covers core functionality but is missing some newer features:

| Feature | PowerShell | C# |
|---------|:---:|:---:|
| Core rip/encode/organize workflow | Yes | Yes |
| Movie mode (Feature file + extras) | Yes | Yes |
| Multi-disc concurrent ripping | Yes | Yes |
| `-Bluray` subtitle fallback | Yes | Yes |
| `-Queue` / `-ProcessQueue` | Yes | Yes |
| Window title management | Yes | Yes |
| Session logging | Yes | Yes |
| `-Documentary` flag | Yes | No |
| `-Tutorial` / `-Fitness` / `-Music` / `-Surf` flags | Yes | No |
| `-Extras` flag (direct output to extras dir) | Yes | No |
| `-StartEpisode` parameter | Yes | No |
| Jellyfin episode naming (`S01E01`) | Yes | No |
| Composite mega-file detection | Yes | No |
| Disc 1 temp dir isolation (`Disc1/`) | Yes | No |
| Series per-disc encoding isolation | Yes | No |
| Empty parent directory cleanup | Yes | No |
| Eject retry with timeout popup | Yes | No |
| Completion fanfare | Yes | No |
| `continue-rip.ps1` resume script | Yes | N/A |
| HandBrake recovery scripts | Yes | No |

## Choosing Between Versions

### Use PowerShell Version If:
- You rip TV series (Jellyfin naming, composite detection, `-StartEpisode`)
- You rip documentaries, tutorials, fitness, music, or surf videos
- You want the latest features
- You want to easily modify the script

### Use C# Version If:
- You only rip movies
- You want a standalone executable
- You prefer statically-typed languages

## Building the C# Version

See [RipDisc/README.md](RipDisc/README.md) for detailed build instructions.

Quick build:
```bash
cd RipDisc
.\build.bat
```

Create self-contained executable:
```bash
cd RipDisc
.\publish.bat
```

## Resuming Failed Rips

If a rip fails after the MakeMKV step, use `continue-rip.ps1` to resume from where it failed:

```powershell
# Continue from HandBrake encoding (step 2)
.\continue-rip.ps1 -title "The Matrix" -FromStep handbrake

# Continue from file organization (step 3)
.\continue-rip.ps1 -title "The Matrix" -FromStep organize

# Continue from open directory (step 4)
.\continue-rip.ps1 -title "The Matrix" -FromStep open
```

### FromStep Options

| Value | Step | Prerequisites |
|-------|------|---------------|
| `handbrake` | 2 | MKV files in `C:\Video\{title}\` |
| `organize` | 3 | MP4 files in output directory |
| `open` | 4 | Output directory exists |

All other parameters work the same as `rip-disc.ps1`:

```powershell
# Resume a TV series rip
.\continue-rip.ps1 -title "Breaking Bad" -Series -Season 1 -FromStep organize

# Resume a Blu-ray rip
.\continue-rip.ps1 -title "Inception" -Bluray -FromStep handbrake

# Resume disc 2 special features
.\continue-rip.ps1 -title "The Dark Knight" -Disc 2 -FromStep handbrake
```

## Additional Tools

- **series-cleanup.ps1** - Utility for cleaning up series naming
- **continue-rip.ps1** - Resume failed rips from a specific step

## Project Structure

```
ripdisc/
├── rip-disc.ps1           # PowerShell implementation
├── continue-rip.ps1       # Resume failed rips from a specific step
├── series-cleanup.ps1     # Series cleanup utility
├── CLAUDE.md              # Development notes
├── README.md              # This file
└── RipDisc/               # C# implementation
    ├── README.md          # C# specific documentation
    ├── build.bat          # Build script
    ├── publish.bat        # Publish script
    └── RipDisc/           # C# project
        ├── Program.cs
        ├── RipDiscApplication.cs
        ├── CommandLineOptions.cs
        ├── ConsoleHelper.cs
        ├── FileHelper.cs
        ├── Logger.cs
        ├── StepTracker.cs
        └── RipDisc.csproj
```

## Contributing

New features are added to the PowerShell scripts first. The C# version should be updated to match when possible.

## License

This project is provided as-is for personal use.

## Notes

- This tool is designed for backing up legally owned physical media
- Ensure you have the legal right to rip any disc you process
- MakeMKV and HandBrake must be properly licensed/installed
