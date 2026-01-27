# RipDisc - C# Console Application

This is a C# port of the PowerShell `rip-disc.ps1` script. It provides the same functionality for automated DVD and Blu-ray disc ripping and encoding using MakeMKV and HandBrake.

## Requirements

- .NET 8.0 or later
- Windows OS
- MakeMKV installed at `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe`
- HandBrake CLI installed at `C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe`

## Building

From the `RipDisc` directory:

```bash
dotnet build -c Release
```

The executable will be located at:
```
RipDisc\bin\Release\net8.0-windows\RipDisc.exe
```

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
- `-drive <string>` - Drive letter (default: D:)
- `-driveIndex <int>` - Drive index for MakeMKV (default: -1)
- `-outputDrive <string>` - Output drive letter (default: E:)

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

All features from the PowerShell script are implemented:

1. **Command-line argument parsing** with validation
2. **4-step processing workflow:**
   - Step 1: MakeMKV ripping to MKV files
   - Step 2: HandBrake encoding to MP4
   - Step 3: File organization (renaming, prefixing, extras folder management)
   - Step 4: Open output directory
3. **Step tracking** with completion summary
4. **Colored console output** matching PowerShell colors
5. **Comprehensive logging** to `C:\Video\logs\`
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

MakeMKV temporary files are stored in:
- Disc 1: `C:\Video\TitleName\`
- Disc 2+: `C:\Video\TitleName\Disc2\`, `C:\Video\TitleName\Disc3\`, etc.

These directories are automatically cleaned up after successful encoding.

## Logs

Session logs are saved to:
```
C:\Video\logs\{title}_disc{disc}_{timestamp}.log
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

The C# version is functionally identical to the PowerShell script with these minor implementation differences:

1. **Process execution**: Uses `System.Diagnostics.Process` instead of PowerShell cmdlets
2. **COM interop**: Uses C# dynamic types for Shell.Application (disc ejection)
3. **File operations**: Uses `System.IO` classes instead of PowerShell file cmdlets
4. **Cross-platform safety**: Uses `[SupportedOSPlatform("windows")]` attributes

## Notes

- The original PowerShell script (`rip-disc.ps1`) remains in the repository root
- Both versions can be used interchangeably
- The C# version may offer better performance and easier distribution as a standalone executable
