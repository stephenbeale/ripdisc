# RipDisc

PowerShell and C# tools for automated DVD and Blu-ray disc ripping using MakeMKV and HandBrake.

## Overview

This repository contains two implementations of the same disc ripping workflow:

1. **PowerShell Script** (`rip-disc.ps1`) - Original implementation
2. **C# Console Application** (`RipDisc/`) - Modern cross-language port

The PowerShell version is the primary implementation and has the most features. The C# version covers core ripping functionality but is behind on some newer features (see [Feature Parity](#feature-parity) below).

## Features

- **Auto-discovery of disc metadata** via MakeMKV + TMDb (title, format, series detection)
- **Automated ripping and encoding** using MakeMKV and HandBrake
- **4-step processing workflow** with progress tracking
- **Movie, TV Series, and genre-based support** (Documentary, Tutorial, Fitness, Music, Surf) with different organization strategies
- **Jellyfin episode naming** for series (`Title-S01E01.mp4`)
- **Composite mega-file detection** skips all-in-one files during series encoding
- **Multi-disc support** with concurrent ripping capability
- **HandBrake queue mode** for sequential encoding after concurrent rips
- **Blu-ray subtitle handling** (scans for forced/foreign-language subs only, burns them in)
- **Real-time MakeMKV progress** streamed to console during rip
- **Feature file identification** (automatically identifies main feature)
- **Extras folder management** for special features
- **Resume failed rips** from any step with `continue-rip.ps1`
- **HandBrake recovery scripts** generated automatically before encoding
- **Corrupt file detection guidance** — diagnose and recover from interrupted rips
- **Comprehensive error handling** with recovery guidance
- **Session logging** for debugging and recovery, with the log path shown as a clickable link at the end of every run
- **Drive readiness checks** before operations
- **Interactive prompts** for confirmation and conflict resolution
- **Window title management** for tracking concurrent operations
- **Console close button protection** prevents accidental window closure
- **Automatic disc ejection** after successful rip

## Auto-Discovery

When `-title` is omitted, the PowerShell script automatically discovers disc metadata:

1. **Reads disc info** via MakeMKV's info mode (disc name, type, title count)
2. **Cleans the disc name** (strips suffixes like `_D1`, `_WS`, replaces underscores, title-cases)
3. **Searches TMDb** (The Movie Database) for the cleaned title
4. **Auto-populates** `-title`, `-Bluray`, `-Series`, `-Season`, and `-Disc` based on results
5. **Prompts for confirmation** — accept, edit, or abort

If `-title` is provided, discovery is skipped (only disc format auto-detection for `-Bluray` runs).

### What Gets Auto-Detected

| Parameter | Auto-detected? | Source |
|-----------|---------------|--------|
| `-title` | Yes | TMDb search, cleaned disc name, or manual fallback |
| `-Bluray` | Yes | MakeMKV disc type (`Blu-ray disc`) |
| `-Series` | Yes | TMDb media type (`tv`) |
| `-Season` | Partial | Regex from disc name (e.g. `S01`, `Season 1`) |
| `-Disc` | Partial | Regex from disc name (e.g. `D2`, `Disc 2`) |
| Genre flags | No | Always manual (`-Documentary`, `-Music`, etc.) |
| `-Extras` | No | Always manual |
| `-StartEpisode` | No | Always manual |

### TMDb API Key Setup

To enable TMDb searching, either:
- Run `setup.ps1` and enter your key when prompted (saved to `ripdisc-config.json`)
- Or set the `TMDB_API_KEY` environment variable:

```powershell
[Environment]::SetEnvironmentVariable("TMDB_API_KEY", "your_api_key_here", "User")
```

Get a free API key at [themoviedb.org/settings/api](https://www.themoviedb.org/settings/api).

Without a TMDb key, the script still works — it uses the cleaned disc name as the title and falls back to manual input if the name is too generic.

## Getting Started

### New to this? Start here:

1. **Download** — Click the green **Code** button above, then **Download ZIP**. Extract it somewhere (e.g. `C:\RipDisc\`).
2. **Double-click `Start.bat`** — This launches the setup wizard, which will:
   - Check for MakeMKV and HandBrakeCLI on your system
   - Help you install anything that's missing
   - Ask you to pick your disc drive and output drive
   - Save your settings so you only do this once
3. **Rip a disc** — Open PowerShell in the RipDisc folder and run:

```powershell
.\rip-disc.ps1 -title "The Matrix"
```

Or let it auto-detect the disc title (requires a free TMDb API key — setup will explain):

```powershell
.\rip-disc.ps1
```

That's it. Insert a disc, run the script, and it handles ripping, encoding, and organising the files.

### Already familiar with PowerShell?

Run `.\setup.ps1` directly instead of the bat file. Or skip setup entirely — the scripts auto-detect tool locations and fall back to sensible defaults.

### C# Version

```bash
cd RipDisc\RipDisc\bin\Release\net8.0-windows
.\RipDisc.exe -title "The Matrix"
```

## Requirements

- **Windows 10/11**
- **[MakeMKV](https://www.makemkv.com/download/)** — reads DVD and Blu-ray discs (setup will help you install it)
- **[HandBrakeCLI](https://handbrake.fr/downloads2.php)** — encodes video files (setup will help you install it)
- **PowerShell 5.1+** (included with Windows 10/11)
- **.NET 8.0+** (only needed for the C# version)

## Configuration

All paths and defaults are stored in `ripdisc-config.json` (created by `setup.ps1`). You can also create it manually from the sample:

```powershell
Copy-Item ripdisc-config.sample.json ripdisc-config.json
```

If no config file exists, the scripts auto-detect tool locations by searching the PATH, Windows registry, and common install directories.

## Usage

Both versions use the same command-line parameters:

```
-title <string>         (Optional) Title of the movie or series (auto-discovered if omitted)
-series                 Flag for TV series
-season <int>           Season number (default: 0)
-disc <int>             Disc number (default: 1)
-drive <string>         Drive letter (default: D:)
-driveIndex <int>       Drive index for MakeMKV (default: -1)
-outputDrive <string>   Output drive letter (default: E:)
-extras                 Flag for extras-only disc
-queue                  Queue encoding instead of running immediately
-bluray                 Blu-ray mode (outputs to Bluray folder, forced subtitle scan)
-documentary            Documentary mode (outputs to Documentaries folder)
-tutorial               Tutorial mode (outputs to Tutorials folder)
-fitness                Fitness mode (outputs to Fitness folder)
-music                  Music mode (outputs to Music folder)
-surf                   Surf mode (outputs to Surf folder)
-startEpisode <int>     Starting episode number for series (default: 1)
```

### Documentary / genre series (multi-disc box sets)

Combine `-series` with a genre flag (`-documentary`, `-tutorial`, `-fitness`, `-music`, `-surf`) for a
multi-disc box set that still belongs under the genre folder rather than `Series\` - for example, a
7-episode documentary spread across 5 discs, where every disc reports the same or a near-identical
disc label so there's no way to tell discs apart automatically. You supply `-disc N` yourself each
time (the same as any other multi-disc rip); the script numbers episodes sequentially and moves them
into a single flat folder, no matter how many episodes end up on each individual disc.

Episode numbering carries across sessions automatically: `-startEpisode` is optional. If you omit it,
the script scans the destination folder for the highest existing `-E##` (or `S##E##`) file and
continues from there - rip disc 1 today, disc 4 next week, and the numbering picks up correctly
without you having to remember or compute where it left off. Pass `-startEpisode` explicitly only if
you need to override that (e.g. re-ripping a disc out of order).

```powershell
# Disc 1 of a 5-disc documentary box set - lands as episodes 1-2
.\rip-disc.ps1 -title "Martin Scorsese Presents the Blues" -documentary -series -disc 1

# Disc 2, ripped in a later session - continues automatically at episode 3
.\rip-disc.ps1 -title "Martin Scorsese Presents the Blues" -documentary -series -disc 2
```

Keep `-title` **identical for every disc in the set**. Continuation works by scanning the
shared destination folder for the highest existing episode number, so a title that varies
per disc sends each one to its own folder and numbering restarts at 1 every time.

#### Episode names

Episodes are titled automatically from the disc's own volume label, normalised to title
case with underscores replaced by spaces - so a disc labelled `WARMING_BY_THE_DEVILS_FIRE`
produces:

```
Martin Scorsese Presents the Blues - S01E04 - Warming By The Devils Fire.mp4
```

This only applies when a disc holds exactly one episode; one label cannot name several
files. Use `-episodeNames` to set them explicitly - it always overrides the disc label, and
is the only option when a disc yields multiple episodes:

```powershell
# Two episodes on one disc, named explicitly
.\rip-disc.ps1 -title "The Blues" -documentary -series -disc 3 `
    -episodeNames "The Road to Memphis", "Warming by the Devil's Fire"
```

Names are matched to files in the order MakeMKV emits them. Any episode without a name
falls back to plain `<title>-E##.mp4` numbering, as does a disc whose label is missing or
generic (`DVD_VIDEO`, `UNTITLED`, and similar).

Two cases where no label is available, so `-episodeNames` is required:

- **`-driveIndex` was used** - the MakeMKV drive list is skipped entirely on that path, and
  with no drive letter there is nothing to ask Windows about either.
- **`continue-rip.ps1`** - it resumes after the disc is done and never reads it.

MakeMKV also leaves the label blank for some drives (reproducibly so on USB DVD units); the
script falls back to querying Windows for the same drive letter before giving up.

### Examples

**Rip a disc with auto-discovery (no title needed):**
```powershell
.\rip-disc.ps1
```

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

**Rip a music disc:**
```powershell
.\rip-disc.ps1 -title "Metallica Live" -music
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

### Documentary / genre series (multi-disc box set, `-documentary -series`)

```
E:\Documentaries\Martin Scorsese Presents the Blues\
├── Martin Scorsese Presents the Blues-E01.mp4    (Disc 1)
├── Martin Scorsese Presents the Blues-E02.mp4    (Disc 1)
├── Martin Scorsese Presents the Blues-E03.mp4    (Disc 2)
├── Martin Scorsese Presents the Blues-E04.mp4    (Disc 3)
└── ...
```

No per-disc subfolders survive - each disc's episodes are numbered and moved into the shared title
folder, and the empty per-disc folder is removed. This is the same layout `-tutorial -series`,
`-fitness -series`, `-music -series` and `-surf -series` produce, just rooted at their own genre
folder (`Tutorials\`, `Fitness\`, `Music\`, `Surf\`). Add `-season N` if the box set genuinely has
seasons, and files use `SeriesName-S01E01.mp4` instead.

### Tutorials

```
E:\Tutorials\TutorialName\
├── TutorialName-Feature.mp4
└── extras\
    └── TutorialName-bonus.mp4
```

### Fitness

```
E:\Fitness\WorkoutName\
├── WorkoutName-Feature.mp4
└── extras\
    └── WorkoutName-bonus.mp4
```

### Music

```
E:\Music\ArtistName\
├── ArtistName-Feature.mp4
└── extras\
    └── ArtistName-behind-the-scenes.mp4
```

### Surf

```
E:\Surf\SurfTitle\
├── SurfTitle-Feature.mp4
└── extras\
    └── SurfTitle-bonus.mp4
```

### Blu-ray

```
F:\Bluray\MovieName\
├── MovieName-Feature.mp4
└── extras\
    ├── MovieName-t01.mp4
    └── MovieName-t02.mp4
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

- Each disc uses an isolated temporary directory (`C:\Video\{title}\Disc1\`, `C:\Video\{title}\Disc2\`, etc.) — even single-disc rips use `Disc1\` so that a concurrent extras rip on Disc 2 won't collide with identically-named MakeMKV output files
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
| Auto-discovery (disc metadata + TMDb) | Yes | No |
| Core rip/encode/organize workflow | Yes | Yes |
| Movie mode (Feature file + extras) | Yes | Yes |
| Multi-disc concurrent ripping | Yes | Yes |
| `-Bluray` output dir + forced subtitle scan | Yes | No |
| `-Queue` / `-ProcessQueue` | Yes | Yes |
| Window title management | Yes | Yes |
| Session logging | Yes | Yes |
| `-Documentary` flag | Yes | No |
| `-Tutorial` / `-Fitness` / `-Music` / `-Surf` flags | Yes | No |
| Genre series (`-Documentary`/etc. combined with `-Series`) | Yes | No |
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

## Recovering from Failures

When a rip fails, the error output tells you exactly which step failed and what to do next. There are two recovery tools available depending on the situation.

### Recovery Scripts (generated automatically)

Every time `rip-disc.ps1` starts encoding, it generates a recovery script at `C:\Video\recovery_{title}_{date}.ps1`. This script contains the exact HandBrake commands needed to encode any remaining MKV files.

**When to use:** HandBrake crashed, was interrupted, or the system lost power during encoding (Step 2). The MKV source files are intact and just need to be re-encoded.

```powershell
# Run the recovery script shown in the error output
cd C:\Video
& '.\recovery_The Matrix_2026-03-29.ps1'
```

The recovery script skips any files that already have a matching `.mp4` in the output directory, so it only encodes what's missing.

**Important:** The recovery script is deleted automatically after a successful encode. If the script doesn't exist, use `continue-rip.ps1` instead (see below).

### continue-rip.ps1 (resume from any step)

Use `continue-rip.ps1` to resume from any step after the initial MakeMKV rip:

```powershell
# Continue from HandBrake encoding (step 2)
.\continue-rip.ps1 -title "The Matrix" -FromStep handbrake

# Continue from file organization (step 3)
.\continue-rip.ps1 -title "The Matrix" -FromStep organize

# Continue from open directory (step 4)
.\continue-rip.ps1 -title "The Matrix" -FromStep open
```

#### FromStep Options

| Value | Step | Prerequisites |
|-------|------|---------------|
| `handbrake` | 2 | MKV files in `C:\Video\{title}\Disc{N}\` |
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

### Handling Partial or Corrupt Files

Both `continue-rip.ps1` and recovery scripts skip files that already have a matching `.mp4`, so you may need to delete partial outputs first:

```powershell
# Delete the corrupt/partial output from the failed encode
Remove-Item "F:\DVDs\Django Unchained\Django Unchained-O1_t00.mp4"

# Then resume — it will re-encode the deleted file and skip the rest
.\continue-rip.ps1 -title "Django Unchained" -FromStep handbrake -OutputDrive F
```

### When Recovery Won't Work (Corrupt MKV)

If the disc was **ejected or removed during the MakeMKV rip** (Step 1), the MKV file may be corrupt even if it appears to be a normal size. Symptoms:

- HandBrake reports: `scan: unrecognized file type` and `found 0 valid title(s)`
- ffprobe reports: `Invalid data found when processing input`

**To check whether your MKV is usable:**

```powershell
# Check the file exists and has a reasonable size (feature film = 4-8 GB typically)
Get-ChildItem 'C:\Video\{title}\Disc1\*.mkv'

# Verify the file is valid (requires ffprobe — included with ffmpeg)
ffprobe 'C:\Video\{title}\Disc1\A1_t00.mkv'
```

If ffprobe shows stream info (duration, codec, resolution), the file is fine — run the recovery script or `continue-rip.ps1`. If ffprobe reports `Invalid data found`, the MKV is corrupt and **you must re-rip from disc**:

```powershell
# Clean up the corrupt files
Remove-Item 'C:\Video\{title}' -Recurse -Force
Remove-Item 'C:\Video\recovery_{title}_*.ps1' -Force

# Re-insert the disc and rip again
.\rip-disc.ps1
```

## Additional Tools

- **series-cleanup.ps1** - Utility for cleaning up series naming
- **continue-rip.ps1** - Resume failed rips from a specific step

## Project Structure

```
ripdisc/
├── Start.bat              # Double-click to get started (launches setup)
├── setup.ps1              # First-run setup (detects/installs tools, creates config)
├── Load-Config.ps1        # Shared config loader (dot-sourced by scripts)
├── ripdisc-config.sample.json  # Sample configuration file
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
