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
- **Episode naming** for series (`Title-S02-E01.mp4`), with extras told apart from episodes (TheDiscDB disc mapping, then TMDb runtimes, then title length), a confirm/edit step, a rename manifest and a one-command undo
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
- **Automatic disc ejection** after successful rip, optional via `-NoEject`
- **Completion fanfare** ([Console]::Beep melody), optional via `-NoSound`
- **eBay sold-price check** — prints a clickable eBay UK sold-listings search URL for the ripped title, optional via `-CheckEbayPrice`
- **Disc type shown per drive** in the drive listing (Audio CD, Blu-ray, DVD-Video, or a size-based guess for a data disc) — a quick check to catch the wrong disc before a rip even starts

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
cd RipDisc\RipDisc.Cli\bin\Release\net8.0-windows
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
-startEpisode <int>     Starting episode number for series (default: 1). For plain -series
                        Disc 2+, you are asked at the start of the run if this is omitted
-noSound                Skip the completion fanfare (Console.Beep melody)
-noDiscDb               Skip the TheDiscDB lookup for plain -series episode/extra naming
-noEject                Skip ejecting the disc after the MakeMKV rip (rip-disc.ps1 only —
                        continue-rip.ps1 accepts it for command-line compatibility but
                        ignores it, since it never runs the rip/eject step)
-checkEbayPrice         Print a clickable eBay UK sold-listings search URL for the title in
                        the FILE SUMMARY (Buy It Now, Very Good+ condition, UK only, sold
                        listings) — a convenience for checking what the physical disc might
                        be worth, not part of the rip itself
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
# or leave -startEpisode off: you're asked before the rip starts, with the next number
# after Disc1's episodes pre-filled (press Enter to accept)
```

### TV series episode naming (`-series`)

At the end of a plain `-series` rip (Step 3), the files in the disc's `DiscN` folder are renamed:

| Before | After |
|--------|-------|
| `title_t00.mp4` | `Silicon Valley-S02-E01.mp4` |
| `title_t01.mkv` | `Silicon Valley-S02-E02.mkv` (extension is never changed) |
| `title_t02.mp4` (a 5-minute featurette) | `..\..\Season 0\Silicon Valley-S02-D1-Extra01.mp4` |

- **No disc number in episode names** - episodes stay in `Season N\DiscN\`, which already says which disc they came from.
- **Extras go in `<Series>\Season 0\`** - one folder at series level, alongside the `Season N` folders,
  shared by every disc of every season. Jellyfin did not pick up the earlier `DiscN\extras\` layout and
  reads a series-level `Season 0` folder as Season 00 (so the name was changed from `Season 0` on 2026-10-06). Because the folder is shared and extras are
  numbered per disc, an extra's name carries its season and disc (`-S02-D1-Extra01`), so two discs'
  `Extra01` never collide. The folder is only created when there are extras. Extras a pre-2026-10-05 run
  left in `DiscN\extras\`, or in `<Series>\Specials\` (the earlier name of `Season 0`, 2026-10-05 runs), still keep their numbers and can still be undone. Existing `Specials` folders are not moved automatically; rename them to `Season 0` by hand if you want them in one place.
- **No `-season`** - the tag falls back to `S01` (`Fargo-S01-E01.mp4`); the folder layout is unchanged (no Season folder).
- **Numbering** starts at `-startEpisode` (default 1) and runs in MakeMKV title order. It does not
  continue across discs on its own: for Disc 2+ without `-startEpisode` you are asked for the starting
  number before the rip begins, with the next number after the earlier discs' episodes pre-filled.
  Disc 1, or an explicit `-startEpisode`, never prompts. A failed run's suggested `continue-rip.ps1`
  command carries the answer, so a resumed rip doesn't ask again.
- **TheDiscDB first** - before the rip, while the disc is still in the drive, the script reads the
  disc's file listing and computes its [TheDiscDB](https://thediscdb.com) content hash (the same
  size-based MD5 TheDiscDB uses: `BDMV\STREAM\*.m2ts` on Blu-ray, `VIDEO_TS\*` on DVD - no file
  contents are read). If TheDiscDB knows the disc, its title-by-title mapping decides which file is
  which episode (published episode numbers, e.g. `E08` on a Disc 2) and which are extras, and the
  confirmation table shows `TheDiscDB` in its Source column. Files are lined up with TheDiscDB's
  titles by duration *in MakeMKV order*, so a different MakeMKV minimum-length setting (which shifts
  the `_tNN` numbers) or the skipped play-all title doesn't throw it off. No API key is needed.
  Extras TheDiscDB names keep that name after the number:
  `30 Rock-S01-D2-Extra01-The C Word Deleted Scene.mp4`.
  No match, offline, or any error: one line is shown and logged, and naming falls back to TMDb and
  the median heuristic below - the lookup never stops a rip (8-second timeout). Turn it off with
  `-noDiscDb`. With `-driveIndex` but no `-drive`, the drive letter isn't known, so it is skipped.
- **Extras vs episodes** - extras get `Extra##` names and no episode number. With a TMDb key, the
  season's published per-episode runtimes are used: each title is compared with the runtime of the
  episode it would become (within 15% or 3 minutes). A title under 60% of that runtime is an extra;
  anything else that doesn't match is kept as an episode but flagged. Without TMDb (no key, no match,
  offline) titles under 60% of the median title length are extras. A "play all" title (about the sum
  of the others) is always an extra. You're warned if numbering would run past the season's last
  episode on TMDb.
- **Confirm before renaming** - the planned names are shown with actual and expected durations
  and where each decision came from (`TheDiscDB`, `TMDb`, `Length`, or `You` after an edit).
  Enter accepts, `n` leaves every file as it is, `e` lets you switch rows between episode and extra.
- **Manifest and undo** - `rename-manifest.csv` (`OriginalName,NewName,Kind,OriginalPath,NewPath,Timestamp`)
  is written in the Disc folder *before* anything is renamed, and `undo-rename.ps1` is copied next to it.
  `NewName` is the path relative to the Disc folder (`..\..\Season 0\<name>` for extras), so undo moves this
  disc's extras back and removes `Season 0` only if that empties it (other discs' extras stay). Renames
  never overwrite an existing file. To undo:

```powershell
& "E:\Series\Silicon Valley\Season 2\Disc1\undo-rename.ps1" -WhatIf   # preview
& "E:\Series\Silicon Valley\Season 2\Disc1\undo-rename.ps1"           # rename back
# or, from the repo:  .\undo-rename.ps1 -ManifestPath "<Disc folder>\rename-manifest.csv"
```

Undo skips (with a warning) files that are missing or whose original name is already taken, and refuses
any manifest row that points anywhere other than the Disc folder, the series-level `Season 0` folder, or
(older manifests) the Disc folder's `extras` subfolder.

- **Non-interactive** - `continue-rip.ps1 -Yes` never stops at the series prompts: on Disc 2+ it takes the
  suggested start episode (TheDiscDB's first episode, else the next after earlier discs, else
  `-startEpisode`/1, flagged as a guess), and it shows the confirmation table and accepts it as planned.
  Both automatic choices are written to the log. `rip-disc.ps1` has no equivalent non-interactive flag.
- **Known limitation: `-queue` / C# `-processQueue`** - a queued rip's encode and organize run in the C#
  queue processor, which does not have any of this: no `S##-E##` renames, extras detection, TheDiscDB
  lookup, `Season 0` folder, manifest or undo. It still writes series files straight into the Season
  folder with the older `Title-S##-...` prefix naming. Porting it has been deferred; for now, rip series
  discs without `-queue` to get this naming.

Genre series (`-series` with `-documentary` etc.) keeps its own naming, described below.

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

**Rip quietly overnight, leave the disc in the drive:**
```powershell
.\rip-disc.ps1 -title "The Matrix" -noSound -noEject
```

**Check what the physical disc might be worth after ripping:**
```powershell
.\rip-disc.ps1 -title "Inception" -bluray -checkEbayPrice
# FILE SUMMARY includes a clickable eBay UK sold-listings search URL
# (Buy It Now, Very Good+ condition, sold listings only)
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
├── Season 0\
│   └── SeriesName-S02-D1-Extra01-Making Of.mp4
└── Season 2\
    ├── Disc1\
    │   ├── SeriesName-S02-E01.mp4
    │   ├── SeriesName-S02-E02.mp4
    │   ├── rename-manifest.csv
    │   └── undo-rename.ps1
    └── Disc2\
        ├── SeriesName-S02-E03.mp4
        └── ...
```

### TV Series (no season)

```
E:\Series\SeriesName\
└── Disc1\
    ├── SeriesName-S01-E01.mp4
    ├── SeriesName-S01-E02.mp4
    └── ...
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
├── MovieName-Feature-BluRay.mp4
└── extras\
    ├── MovieName-t01.mp4
    └── MovieName-t02.mp4
```

With `-Bluray`, the main feature is named `<Title>-Feature-BluRay.<ext>` so Jellyfin can tell a Blu-ray rip from a DVD of the same film.

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

Before anything else runs, both scripts validate the output path and check the destination drive is actually connected (via `Test-DriveReady`) — a missing or disconnected `-OutputDrive` stops the script immediately with a clear message, rather than failing later mid-encode or offering to continue with a broken path.

If an error occurs:
- Window title shows `-ERROR` suffix
- Completed steps are displayed in green
- Remaining steps are listed with manual instructions
- **If the MakeMKV rip itself already completed**, a ready-to-paste `continue-rip.ps1` command is printed under `--- RETRY WITH continue-rip.ps1 ---`, built from this run's own inputs (title, series/season/disc, genre flags, `-StartEpisode`, `-EpisodeNames`, `-NoSound`, and for series the TheDiscDB disc hash as `-DiscDbHash` or `-NoDiscDb`) with `-FromStep` set to whichever step failed (`handbrake`, `organize`, or `open`) — copy it as-is to resume without re-ripping the disc
  - Not shown when Step 1 (the rip) itself failed — `continue-rip.ps1` has no ripped MKV files to resume from in that case, so re-running `rip-disc.ps1` is the only option
- Relevant directory is opened for inspection
- Log file location is provided

**Drive lookup (`rip-disc.ps1` only):** before ripping starts, the script queries MakeMKV directly (`Looking up drive X in MakeMKV...`) to map your drive letter to its internal index. This query is bounded by a 60-second timeout — a malfunctioning or slow-to-spin-up drive can legitimately take 30+ seconds here, so this isn't necessarily a hang. If it does time out, the error explains why no log file exists yet at this point, checks for and reports any leftover `makemkvcon`/`makemkvcon64` process from a previous Ctrl+C'd run (which can hold the drive exclusively), and lists concrete retry options, including `-DriveIndex` to skip the lookup entirely.

**Disc type shown in the drive listing:** every non-busy drive in the `MakeMKV drives:` listing shows what's actually in it — `Audio CD`, `Blu-ray`, `DVD-Video`, or a size-based guess for a plain data disc (`Data disc (CD-ROM-sized)` etc.) — worked out independently of MakeMKV (volume label, `BDMV`/`VIDEO_TS` folder presence, disc capacity). This is a quick visual check to catch the wrong disc in the drive (e.g. a music CD instead of the movie DVD) before a rip even starts, not just after MakeMKV fails on it — if `rip-disc.ps1` does fail with a SCSI `ILLEGAL MODE FOR THIS TRACK` error and the disc turns out to be an audio CD, the error message says so directly and points at `ripaudio`'s `rip-audio.ps1` instead.

**Failure classification:** a MakeMKV run that reads the disc successfully but then fails to save any titles (e.g. the drive disconnects mid-rip) is reported as a drive disconnect, not as "no disc detected" — the two are distinguished by the specific Windows error text rather than a generic "0 titles" match, so a real hardware disconnect isn't confused with an empty/unreadable drive.

## Feature Parity

The PowerShell scripts are the primary implementation. The C# version covers core functionality but is missing some newer features:

| Feature | PowerShell | C# |
|---------|:---:|:---:|
| Auto-discovery (disc metadata + TMDb) | Yes | No |
| Core rip/encode/organize workflow | Yes | Yes |
| Movie mode (Feature file + extras) | Yes | Yes |
| Multi-disc concurrent ripping | Yes | Yes |
| `-Bluray` output dir + forced subtitle scan | Yes | No |
| `-Queue` / `-ProcessQueue` | Yes | Yes (queued series rips don't get episode naming - see TV series episode naming) |
| Window title management | Yes | Yes |
| Session logging | Yes | Yes |
| `-Documentary` flag | Yes | No |
| `-Tutorial` / `-Fitness` / `-Music` / `-Surf` flags | Yes | No |
| Genre series (`-Documentary`/etc. combined with `-Series`) | Yes | No |
| `-Extras` flag (direct output to extras dir) | Yes | No |
| `-StartEpisode` parameter | Yes | No |
| Episode naming (`S02-E01`), extras detection, rename manifest + undo | Yes | No |
| TheDiscDB disc lookup for series naming (`-NoDiscDb`) | Yes | No |
| Composite mega-file detection | Yes | No |
| Disc 1 temp dir isolation (`Disc1/`) | Yes | No |
| Series per-disc encoding isolation | Yes | No |
| Empty parent directory cleanup | Yes | No |
| Eject retry with timeout popup | Yes | No |
| Completion fanfare | Yes | No |
| `-NoSound` / `-NoEject` flags | Yes | No |
| Suggested `continue-rip.ps1` retry command on failure | Yes | No |
| Live disc-label lookup for the target drive (avoids stale cached name) | Yes | No |
| Bounded drive-query timeout (60s) with leftover-process diagnostics | Yes | No |
| MakeMKV rip process hang safety net (kills a stuck/unresponsive process rather than blocking forever) | Yes (stuck-sector detection) | Yes (flat 4h timeout) |
| `-CheckEbayPrice` (eBay UK sold-listings search URL in the FILE SUMMARY) | Yes (`rip-disc.ps1` only) | No |
| Disc type shown per-drive in the drive listing (Audio CD/Blu-ray/DVD-Video/data disc) | Yes (`rip-disc.ps1` only) | No |
| Device-disconnect vs. no-disc error classification | Yes | Yes |
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

Use `continue-rip.ps1` to resume from any step after the initial MakeMKV rip. If `rip-disc.ps1`
fails after the rip completes, it now prints the exact command to run — see
[Error Handling](#error-handling) — so you usually won't need to build one of these by hand:

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

**`-Yes`** (`continue-rip.ps1` only) skips confirmation prompts for a plain `-Series` resume: the
Disc 2+ starting-episode prompt and the pre-rename confirmation table are both answered
automatically (suggested start episode, plan accepted as shown — see
[TV series episode naming](#tv-series-episode-naming--series) above for exactly what each one
defaults to). Both automatic choices are written to the log. `rip-disc.ps1` has no equivalent flag.

All other parameters work the same as `rip-disc.ps1`:

```powershell
# Resume a TV series rip
.\continue-rip.ps1 -title "Breaking Bad" -Series -Season 1 -FromStep organize

# ...with the TheDiscDB hash from the original rip's retry command or log, so episodes
# and extras are still named from TheDiscDB (this script never reads the disc itself)
.\continue-rip.ps1 -title "30 Rock" -Series -Season 1 -Disc 2 -FromStep organize -DiscDbHash 16E974A41F04B04E0FC0F7B27EA29758

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
- **rename-series.ps1** - Renames series that are ALREADY ripped on disk; dry run by default, `-Apply` to rename. Full usage in [rename-series.ps1](#rename-seriesps1-rename-series-already-ripped-on-disk) below.
- **continue-rip.ps1** - Resume failed rips from a specific step
- **undo-rename.ps1** - Reverses a plain `-Series` rename using the Disc folder's `rename-manifest.csv` (see [TV series episode naming](#tv-series-episode-naming--series) above). A copy is placed in every Disc folder automatically; the repo-root copy takes an explicit `-ManifestPath`
- **SeriesRetroRename.ps1** - Folder discovery and numbering across discs for `rename-series.ps1`, dot-sourced, not run directly
- **SeriesEpisodes.ps1** - Shared plain `-Series` naming logic (episode/extras classification, TheDiscDB lookup, manifest writing), dot-sourced by both `rip-disc.ps1` and `continue-rip.ps1` rather than duplicated — not run directly

### rename-series.ps1 (rename series already ripped on disk)

The retroactive twin of the `-Series` rip naming. Point it at a series folder, a Season folder or a single
Disc folder. It uses the same names, episode/extra detection and series-level `Season 0` folder as a live rip.
TheDiscDB is not used (it needs the physical disc's hash); TMDb runtimes (or, without a key, title lengths)
decide episode vs extra.

```powershell
.\rename-series.ps1 "F:\Series\Silicon Valley"            # dry run (default): prints each folder's plan, changes nothing
.\rename-series.ps1 "F:\Series\Silicon Valley" -WhatIf    # also a dry run, always
.\rename-series.ps1 "F:\Series\Silicon Valley" -Apply     # asks per folder before renaming
.\rename-series.ps1 "F:\Series\Silicon Valley" -Apply -Yes   # accepts every folder's plan without asking
```

- **`-StartEpisode N`** - episode number for the first folder of each season. An explicit value always wins,
  even over the start suggested from earlier discs. Without it, numbering continues across discs.
- **`-MarkSpecial` / `-MarkExtra` / `-MarkEpisode`** - force matching files (wildcards on the ORIGINAL file
  names) to a kind. Specials become `<Title>-S00-E##` in `Season 0`; extras become
  `<Title>-S##-Extra##` (or `-D#-Extra##`) in `Season 0`; episodes stay `S##-E##`.
- **Prompt edit codes** - at a folder's confirmation prompt, Enter accepts, `n` declines (so does end-of-input),
  `e` edits. In the edit, enter a row number followed by a code: `2s` = special, `2x` = extra, `2e` = episode.
- **Flagged, not converted** - a title at least 1.6x the expected length is flagged "a special?" but stays an
  episode, because double episodes look the same by length. Use `-MarkSpecial` or `2s` to move it.
- **`Season 0` folder** - extras and specials of all seasons go to one `<Series>\Season 0\` folder, which
  Jellyfin reads as Season 00. It is never itself treated as a season to rename.
- **Layouts understood** - `Series N`, `Season N`, `<Title>-Series N`, `Disc N` (with a space), `<Show>-Disc N`
  folders beside a Season folder (they join that season), and several discs in one folder told apart by
  `Disc N` in the file names. Files sort numerically (`(2)` before `(10)`).
- **Manifest and undo** - every folder gets `rename-manifest.csv` + `undo-rename.ps1` BEFORE anything moves;
  nothing is overwritten. Running a folder's `undo-rename.ps1` reverses it, and when every row was undone the
  manifest is retired as `rename-manifest.undone-<stamp>[-n].csv`, so the next apply starts a fresh manifest
  instead of replaying old rows. Undo still accepts manifests recorded with `Specials\` or `DiscN\extras\` paths.

**Example: Boys From the Black Stuff (2026-10-06).** Disc 1 sat in `Season 1` with the 1980 Play for Today
"The Black Stuff" (1:42:26) and a 4:27 extra; Discs 2 and 3 were in `Boys from the Black Stuff-Disc 2` and
`-Disc 3`. Episode 3 had never been ripped, so Disc 3 starts at E04.

```powershell
$s = "F:\Series\Boys From the Black Stuff"
# Disc 1: the long title is a special (S00-E01); the 4:27 title is detected as an extra
.\rename-series.ps1 "$s\Season 1" -MarkSpecial "*Disc 1 - E01*" -Apply
.\rename-series.ps1 "$s\Boys from the Black Stuff-Disc 2" -StartEpisode 1 -Apply   # S01-E01, E02
.\rename-series.ps1 "$s\Boys from the Black Stuff-Disc 3" -StartEpisode 4 -Apply   # S01-E04, E05
# Always run each without -Apply first to check the plan. Re-rip E03 later with -StartEpisode 3.
```

Result: `Season 0\Boys From the Black Stuff-S00-E01.mp4` and `...-S01-Extra01.mp4`; episodes renamed in their
folders (moving them into `Season 1` is a separate manual step). See `Get-Help .\rename-series.ps1 -Full`.

## Project Structure

```
ripdisc/
├── Start.bat              # Double-click to get started (launches setup)
├── setup.ps1              # First-run setup (detects/installs tools, creates config)
├── Load-Config.ps1        # Shared config loader (dot-sourced by scripts)
├── ripdisc-config.sample.json  # Sample configuration file
├── rip-disc.ps1           # PowerShell implementation
├── continue-rip.ps1       # Resume failed rips from a specific step
├── SeriesEpisodes.ps1     # Shared -Series naming logic (dot-sourced, not run directly)
├── rename-series.ps1      # Retroactive series rename of already-ripped folders (dry run by default)
├── SeriesRetroRename.ps1  # Folder discovery/orchestration for rename-series.ps1 (dot-sourced)
├── undo-rename.ps1        # Reverses a -Series rename from a rename-manifest.csv
├── series-cleanup.ps1     # Series cleanup utility
├── CHANGELOG.md           # Release history
├── Roadmap.md             # Open feature requests and backlog
├── CLAUDE.md              # Development notes
├── README.md              # This file
└── RipDisc/               # C# implementation
    ├── README.md          # C# specific documentation
    ├── RipDisc.sln        # Solution (Core, Cli, Tests)
    ├── build.bat          # Build script
    ├── publish.bat        # Publish script
    ├── RipDisc.Core/      # Pipeline, config, queue - no console code (IRipUI seam)
    ├── RipDisc.Cli/       # Console app, builds RipDisc.exe
    └── RipDisc.Tests/     # xUnit tests (dotnet test RipDisc.sln)
```

## Contributing

New features are added to the PowerShell scripts first. The C# version should be updated to match when possible.

## License

This project is provided as-is for personal use.

## Notes

- This tool is designed for backing up legally owned physical media
- Ensure you have the legal right to rip any disc you process
- MakeMKV and HandBrake must be properly licensed/installed
