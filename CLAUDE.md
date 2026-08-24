# RipDisc Project

PowerShell scripts for automated DVD and Blu-ray disc ripping using MakeMKV and HandBrake.

## Git Workflow

When the user says **"make a workflow"**, execute the full git lifecycle. The workflow is **not complete until the PR is approved and merged**:

1. **Branch** - Create a feature branch from main (`feature/<issue-number>-<description>` or `feature/<description>`)
2. **Commit** - Stage and commit all relevant changes with a conventional commit message
3. **Push** - Push the branch to origin (`git push -u origin <branch>`)
4. **PR** - Create a pull request via `gh pr create` with summary and test plan
5. **Approve PR** - Approve via `gh pr review --approve`, then merge via `gh pr merge --squash --delete-branch`
6. **Return to main** - `git checkout main && git pull`

## Session Notes

### 2026-01-19 - Multi-Disc Concurrent Ripping Implementation

**Work Completed:**

**PR #6 - Configurable Output Drive**
- Added `-OutputDrive` parameter to make output drive configurable (default E:)
- Accepts both "E" and "E:" formats
- Allows users to change output location without modifying script

**PR #7 - Concurrent Disc Ripping Support**
- Implemented disc-specific MakeMKV temp directories:
  - Disc 1: `C:\Video\$title\`
  - Disc 2+: `C:\Video\$title\Disc$Disc\`
- Added "Special Features" naming convention for Disc 2+ files: `MovieName-Special Features-filename.mp4`
- Set window titles to show film name for easy identification of concurrent rip operations
- Added status suffixes to window titles:
  - `-INPUT` (waiting for user confirmation)
  - `-ERROR` (failed)
  - `-DONE` (completed successfully)
- Enables parallel ripping of multi-disc films using separate terminal windows on different drives

**PR #8 - Extras Disc Window Title Format**
- Changed extras disc window title to use lowercase `-extras` format
- Example: `Die Another Day-extras` instead of `Die Another Day - Disc 2`
- Improved visual consistency with status indicators

**Documentation Updates:**
- Updated `.claude\agents\disc-ripper.md` throughout to reflect all new features
- Added concurrent ripping workflow examples
- Documented window title conventions

**All PRs merged to main branch successfully.**

**Work In Progress:**
- None - all features completed and merged

**Next Steps:**
- Monitor for any issues with concurrent ripping workflow
- Consider future enhancements:
  - Progress tracking for concurrent operations
  - Notification when all concurrent rips complete
  - Support for 3+ disc concurrent operations

**Technical Notes:**
- Concurrent ripping requires separate temp directories to avoid file conflicts
- Window title changes help users track multiple concurrent rip operations
- Special Features naming convention prevents confusion between main feature and extras
- All changes are backward compatible with existing single-disc workflows
- Drive readiness check (from previous PR) ensures destination drive is available before starting

---

### 2026-01-28 - HandBrake Queue for Sequential Encoding

**Problem:**
When ripping multiple discs concurrently, each session spawns its own HandBrakeCLI process. Multiple concurrent HandBrake workers cause significant CPU contention and slowdown.

**Solution:**
Added queue mode that defers encoding to a separate sequential processing step.

**PR #17 - HandBrake Encoding Queue**

New command-line flags:
- `-queue` — After MakeMKV rip, write encoding job to shared queue file instead of running HandBrakeCLI inline (skips Steps 2-4)
- `-processQueue` — Process all queued encoding jobs sequentially through a single HandBrakeCLI instance

**Usage workflow:**
```powershell
# Rip multiple discs concurrently, each queuing its encode:
RipDisc -title "The Matrix" -queue                      # Terminal 1
RipDisc -title "The Matrix" -disc 2 -queue -driveIndex 1  # Terminal 2

# After all rips complete, encode everything one at a time:
RipDisc -processQueue
```

**Implementation details:**
- Queue file: `C:\Video\handbrake-queue.json`
- File locking protects concurrent writes from parallel rip sessions
- Queue re-read after each completed job to pick up new entries added during processing
- Failed jobs preserved in queue for retry
- Records actual MakeMKV output directory (handles suffixed directory edge case)
- `-queue` and `-processQueue` validated as mutually exclusive
- Both C# and PowerShell implementations updated

**Files changed:**
- `RipDisc/RipDisc/CommandLineOptions.cs` — Added `Queue` and `ProcessQueue` properties
- `RipDisc/RipDisc/RipDiscApplication.cs` — Added `WriteToQueue()`, `RunFromQueue()`, `ProcessAllQueued()`, `QueueEntry` class
- `RipDisc/RipDisc/Program.cs` — Routing for `-processQueue` mode, updated usage text
- `rip-disc.ps1` — Added `-Queue` parameter and queue writing logic

**Technical Notes:**
- `QueueEntry` stores `MakeMkvOutputDir` to handle cases where user chose suffixed directory during Step 1
- `ProcessAllQueued()` uses while-loop with queue re-read to handle concurrent additions
- Window title shows `QUEUED` status when job is added to queue
- Normal mode (without `-queue`) unchanged — fully backward compatible

---

### 2026-01-29 - Blu-ray Subtitle Skip Option

**Problem:**
Blu-ray discs use PGS (Presentation Graphics Stream) subtitles which are image-based. These don't work properly in MP4 containers — they're either not displayed by players or cause playback issues. DVD subtitles (VOB-based) work fine.

**Solution:**
Added `-bluray` flag that skips subtitle extraction entirely for Blu-ray discs.

**PR #18 - Blu-ray Skip Subtitles**

New command-line flag:
- `-bluray` — Skip subtitles during HandBrake encoding (omits `--all-subtitles` and `--subtitle-burned=none`)

**Usage examples:**
```powershell
# Standard Blu-ray rip (no subtitles)
RipDisc -title "Inception" -bluray

# Blu-ray with queue mode
RipDisc -title "Inception" -bluray -queue

# Multi-disc Blu-ray concurrent ripping
RipDisc -title "The Dark Knight" -bluray -queue                    # Terminal 1
RipDisc -title "The Dark Knight" -bluray -disc 2 -queue -driveIndex 1  # Terminal 2
RipDisc -processQueue                                               # After rips complete

# DVD rip (subtitles included by default)
RipDisc -title "Old Movie"
```

**Implementation details:**
- Flag preserved through queue system (`QueueEntry.Bluray` property)
- Works with both C# and PowerShell implementations
- Default behavior unchanged — DVDs still include all subtitles

**Files changed:**
- `RipDisc/RipDisc/CommandLineOptions.cs` — Added `Bluray` property and parser case
- `RipDisc/RipDisc/RipDiscApplication.cs` — Conditional subtitle args, updated `QueueEntry` class
- `RipDisc/RipDisc/Program.cs` — Updated usage text
- `rip-disc.ps1` — Added `-Bluray` parameter and conditional subtitle handling

---

### 2026-02-01 - Bluray Subtitle Fallback

**Problem:**
The previous Bluray implementation completely skipped subtitles. This meant Bluray rips never got subtitles, even when they might work.

**Solution:**
Changed to a "try subtitles first, fallback without" approach for Bluray discs.

**Implementation details:**
- All encodes now try with subtitles first (`--all-subtitles --subtitle-burned=none`)
- For Bluray: if encoding fails, retry without subtitle arguments (PGS incompatibility fallback)
- Subtitles are never burned in — kept as separate streams when possible
- Logs when fallback occurs for troubleshooting

**Files changed:**
- `RipDisc/RipDisc/RipDiscApplication.cs` — Added subtitle fallback logic
- `rip-disc.ps1` — Added subtitle fallback logic
- `continue-rip.ps1` — Added subtitle fallback logic

---

### 2026-02-01 - Continue From Step Script

**Problem:**
When a rip fails at step 2 (HandBrake), step 3 (organize), or step 4 (open), there was no easy way to resume from that point. Users had to manually run the remaining steps or re-rip the entire disc.

**Solution:**
Added `continue-rip.ps1` script that allows resuming from any step after the initial MakeMKV rip.

**PR #21 - Continue From Step Script**

New script: `continue-rip.ps1`

**Parameters:**
- `-FromStep` (required) — Which step to continue from: `handbrake`, `organize`, or `open`
- All other parameters same as `rip-disc.ps1`: `-title`, `-Series`, `-Season`, `-Disc`, `-OutputDrive`, `-Extras`, `-Bluray`, `-Documentary`

**Step mapping:**
| FromStep | Step # | Prerequisites |
|----------|--------|---------------|
| `handbrake` | 2 | MKV files in MakeMKV output directory |
| `organize` | 3 | MP4 files in final output directory |
| `open` | 4 | Final output directory exists |

**Usage examples:**
```powershell
# Continue from HandBrake encoding (step 2) - MKV files must exist
.\continue-rip.ps1 -title "The Matrix" -FromStep handbrake

# Continue from file organization (step 3) - MP4 files must exist
.\continue-rip.ps1 -title "Fargo" -Series -Season 1 -FromStep organize

# Continue from open directory (step 4)
.\continue-rip.ps1 -title "Inception" -FromStep open -Bluray

# Continue with special features disc
.\continue-rip.ps1 -title "The Dark Knight" -Disc 2 -FromStep handbrake
```

**Implementation details:**
- Validates prerequisites exist before starting (MKV files for handbrake, MP4 files for organize)
- Marks skipped steps as "completed" in the step tracker
- Window title shows "CONTINUE" to distinguish from normal rips
- Separate log file with `_continue_` suffix: `{title}_{disc}_continue_{timestamp}.log`
- Same file organization logic as `rip-disc.ps1` (movie/series/documentary modes)
- Same error handling with recovery guidance

**Files added:**
- `continue-rip.ps1` — New script (729 lines)

---

### 2026-02-09 - Fix Disc 1 Concurrent Cleanup & File Lock Logging

**Problem:**
When running concurrent disc rips (e.g. Disc 1 and Disc 2 in separate terminal tabs), Disc 1's cleanup step would try to delete the entire `C:\Video\$title` directory recursively. This nuked `Disc2/` and `Disc3/` subdirectories that were still in use by concurrent rip sessions, causing file lock errors. Additionally, the "File locked" retry messages showed no filename because `$_` inside `catch` blocks referred to the error record, not the pipeline file item.

**Solution:**

**PR #24 - Fix Disc 1 Concurrent Cleanup**

**Bug 1 — Disc 1 temp directory collision:**
- Previously Disc 1 used `C:\Video\$title` (no subdirectory) while Disc 2+ used `C:\Video\$title\Disc$Disc`
- Changed so ALL discs use subdirectories: `C:\Video\$title\Disc$Disc` (e.g. `Disc1`, `Disc2`, `Disc3`)
- Each concurrent rip's temp directory is now isolated, so cleanup only removes its own files

**Bug 2 — `$_` clobbering in ForEach-Object catch blocks:**
- Inside `ForEach-Object` with `try/catch`, PowerShell's `$_` in the `catch` block refers to the error record, not the pipeline file
- Added `$file = $_` at the top of each `ForEach-Object` block and used `$file` throughout
- Retry messages now correctly display the locked filename

**Files changed:**
- `rip-disc.ps1` — Both fixes applied (MakeMKV output dir + 4 rename blocks)
- `continue-rip.ps1` — Both fixes applied (MakeMKV output dir + 3 rename blocks)

---

### 2026-02-16 - Series Episode Renaming & Composite File Exclusion

**Problem:**
When ripping TV series discs, MakeMKV often produces individual episode files plus one composite mega-file containing all episodes concatenated. The script was encoding all files (wasting hours on the composite) and only prefixing filenames with the series title instead of numbering episodes.

**Solution:**

**PR #31 - Series Episode Renaming & Composite File Exclusion**

Two changes:

**1. Composite mega-file detection (Step 2 — HandBrake encoding):**
- In series mode only, if there are 3+ MKV files and the largest is at least 2x the size of the second-largest, it's treated as the composite
- The composite file is excluded from HandBrake encoding (skipped, not deleted)
- The MKV stays on disk but gets cleaned up when the temp directory is deleted after encoding
- If no file meets the threshold, all files are encoded (safe fallback)
- Threshold of 3+ files avoids false positives on 2-episode discs

**2. Jellyfin episode renaming (Step 3 — Organize):**
- Series files renamed from `title_t00.mp4` to `Title-S01E01.mp4` (Jellyfin naming convention)
- Files sorted by name to preserve MakeMKV title order as episode order
- Season tag included when `-Season` is specified, omitted otherwise

**Naming examples:**
| Scenario | Input | Output |
|----------|-------|--------|
| Season 1, file 1 | `title_t00.mp4` | `Fargo-S01E01.mp4` |
| Season 1, file 2 | `title_t01.mp4` | `Fargo-S01E02.mp4` |
| No season, file 1 | `title_t00.mp4` | `Fargo-E01.mp4` |

**Files changed:**
- `rip-disc.ps1` — Composite detection + Jellyfin rename
- `continue-rip.ps1` — Same changes

---

### 2026-02-16 - StartEpisode Parameter for Multi-Disc Seasons

**Problem:**
Episode numbering always started at E01 per disc. Multi-disc seasons (e.g. episodes 1-4 on Disc 1, episodes 5-8 on Disc 2) would produce duplicate episode numbers.

**Solution:**

**PR #32 - Add -StartEpisode Parameter**

New parameter:
- `-StartEpisode` (int, default 1) — Starting episode number for Jellyfin renaming

**Usage examples:**
```powershell
# Disc 1: episodes 1-4 (default, starts at E01)
.\rip-disc.ps1 -title "Fargo" -Series -Season 1 -Disc 1

# Disc 2: episodes 5-8 (starts at E05)
.\rip-disc.ps1 -title "Fargo" -Series -Season 1 -Disc 2 -StartEpisode 5

# Continue script also supports it
.\continue-rip.ps1 -title "Fargo" -Series -Season 1 -FromStep organize -StartEpisode 5
```

**Files changed:**
- `rip-disc.ps1` — Added `-StartEpisode` parameter, used as initial `$episodeNum`
- `continue-rip.ps1` — Same changes

---

### 2026-02-16 - Documentation Updates & Changelog Creation

**Summary:**
Major documentation update to accurately reflect the current state of both PowerShell and C# implementations. Created feature parity table and comprehensive changelog.

**Work Completed:**

**PR #33 - Add Session Notes for Episode Rename Features**
- Added session notes for PR #31 (Series Episode Renaming & Composite File Exclusion)
- Added session notes for PR #32 (StartEpisode Parameter)
- Merged: 2026-02-16T11:57:52Z

**PR #34 - Update README and Add CHANGELOG**
- Updated README Feature List to include all missing features:
  - Composite mega-file detection
  - Jellyfin episode renaming
  - `-StartEpisode` parameter
  - `-Documentary` mode
  - `-Extras` parameter
  - Disc 1 isolation fix
- Created comprehensive CHANGELOG.md documenting all releases and PRs
- Merged: 2026-02-16T12:04:00Z

**PR #35 - Update README to Clarify C# Feature Parity**
- Replaced feature list with Feature Parity Table showing PowerShell vs C# implementation status
- Clearly indicates which features are missing from C# implementation:
  - Composite detection (PS only)
  - Jellyfin episode renaming (PS only)
  - `-StartEpisode` parameter (PS only)
  - `-Documentary` mode (PS only)
  - `-Extras` parameter (PS only)
  - Disc 1 directory isolation (PS only)
- Honest assessment that C# needs porting work to achieve parity
- Merged: 2026-02-16T12:22:22Z

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation:
  - Composite mega-file detection (Step 2 encoding)
  - Jellyfin episode renaming format (`Title-S01E01.mp4`)
  - `-StartEpisode` parameter for multi-disc offset
  - `-Documentary` parameter for genre-based organization
  - `-Extras` parameter for special features mode
  - Disc 1 directory isolation (`Disc1/` subdirectory)
- The README Feature Parity table tracks exactly what's missing

**Technical Notes:**
- PowerShell implementation is feature-complete and production-ready
- C# implementation has core functionality but lacks recent enhancements
- Feature parity table provides clear roadmap for C# porting work
- CHANGELOG.md follows Keep a Changelog format

---

### 2026-02-16 - Eject Retry, Timeout Popup & Completion Fanfare

**Problem:**
Disc eject was timing out intermittently via the COM `Shell.Application` interface — likely due to drive busy states, handle locks, or firmware delays. The timeout handling added in PR #28 caught it, but a single attempt wasn't resilient enough. Users also had no out-of-terminal notification when eject failed, and no audible signal when a rip completed.

**Solution:**

**PR #36 - Add Eject Retry on Timeout**
- Wrapped disc eject in a retry loop (max 2 attempts)
- 2-second delay between attempts to let the drive settle
- Only falls back to "please eject manually" after both attempts fail
- Merged: 2026-02-16

**PR #37 - Add Eject Timeout Popup and Completion Fanfare**

Two additions:

**1. Windows dialog popup on eject timeout:**
- When both eject attempts fail, shows a `System.Windows.Forms.MessageBox` dialog
- Dialog includes film title and drive letter: *"Disc eject timed out for 'Title' on drive D:. It is safe to eject the disc manually."*
- Visible outside PowerShell so user is notified even when in another application

**2. Triumphant completion fanfare:**
- Plays a C major arpeggio melody via `[Console]::Beep`: C5-E5-G5-C6 (pause) G5-C6
- Fires on both normal completion (after Step 4) and queue-mode completion (after "QUEUED!")
- Audible from another room as a distinctive completion signal

**PR #38 - Add Completion Fanfare to continue-rip.ps1**
- Same C major arpeggio melody added to `continue-rip.ps1` for parity with `rip-disc.ps1`
- Merged: 2026-02-16

**Files changed:**
- `rip-disc.ps1` — Eject retry loop, MessageBox popup, fanfare in both completion paths
- `continue-rip.ps1` — Completion fanfare

---

### 2026-02-18 - New Disc Types, Series Fix, Extras Improvements, Cleanup

**PR #39 - Clean Up Empty Parent Directories After Temp Removal**
- After removing MakeMKV temp dir, walk up parent chain deleting empty directories
- Stops at `C:\Video` to avoid removing root working directory
- Applied to both scripts

**PR #40 - Add -Tutorial and -Fitness Disc Type Switches**
- `-Tutorial` outputs to `E:\Tutorials\<title>\`
- `-Fitness` outputs to `E:\Fitness\<title>\`
- Both work identically to `-Documentary` (genre-based folder routing)
- Applied to both scripts including logging and content type labels

**PR #41 - Fix Series Concurrent Disc Rename Conflicts**
- Series encoding now uses per-disc subdirs (`Season 1\Disc1\`, `Season 1\Disc2\`)
- After Step 3 rename, files move up to season folder with file lock retries
- Empty disc subdirs cleaned up after move
- Same pattern as MakeMKV temp dir fix in PR #24

**PR #42 - Improve Extras Disc Renaming**
- Extras files prefixed with title only (`Platoon-title_t00.mp4`)
- No `-extras` or `-Special Features` in filenames
- Lock retries on rename, `-1` suffix on conflicts via `Get-UniqueFilePath`

**PR #43 - Route Extras Disc Output Directly to Extras Subdirectory**
- When `-Extras` is set, HandBrake encodes directly into `<title>\extras\`
- Step 3 skips the move since files are already in the right place
- Works with all genre types (Documentary, Tutorial, Fitness, Movie)

**Files changed:**
- `rip-disc.ps1` — All changes above
- `continue-rip.ps1` — All changes above

**PR #44 - Add -Surf Disc Type Switch**
- `-Surf` outputs to `E:\Surf\<title>\`
- Same pattern as Documentary/Tutorial/Fitness

**Fix: UTF-8 BOM for PowerShell 5.1 Compatibility**
- Scripts saved without BOM caused PowerShell 5.1 to default to Windows-1252 encoding
- Complex nested string expressions like `$([math]::Round($f.Length/1GB, 2))` failed to parse
- Adding UTF-8 BOM (3-byte `EF BB BF` prefix) fixes PS 5.1 while remaining PS 7+ compatible
- **Important:** Always ensure scripts are saved with UTF-8 BOM for PS 5.1 compatibility

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation (see Feature Parity table in README)

---

### 2026-02-22 - Music Switch & Disc 1 Extras Collision Fix

**PR #45 - Add -Music Switch and Fix Disc 1 Extras Move Collision**

Two changes:

**1. `-Music` disc type switch:**
- `-Music` outputs to `<OutputDrive>:\Music\<title>\`
- Same pattern as Documentary/Tutorial/Fitness/Surf (genre-based folder routing)
- Works with `-Extras` for `Music\<title>\extras\`
- Applied to both scripts including parameter declarations, output directory routing, content type labels, genre labels, and logging

**2. Disc 1 non-feature move collision fix:**
- When running Disc 1 and Extras disc concurrently, both MakeMKV runs produce identically-named files (`title_t00`, `title_t01`, etc.)
- Disc 1's Step 3 moves non-feature files to `extras/`, but the Extras disc may have already placed identically-named files there
- Previously used bare `Move-Item -Destination extras -ErrorAction SilentlyContinue` which silently failed on name collisions, leaving files stranded in the main directory
- Now uses `Get-UniqueFilePath` with `-1` suffix handling and verbose output, matching what the Disc 2+ path already did

**Files changed:**
- `rip-disc.ps1` — Music switch + collision fix
- `continue-rip.ps1` — Music switch + collision fix

**PR #46 - Session Notes for PR #45**
- Added 2026-02-22 session notes to CLAUDE.md

**PR #47 - Update CHANGELOG**
- Added 2026-02-22 section to CHANGELOG.md documenting Music switch and collision fix

**PR #48 - Update README Feature Parity & Usage**
- Added `-Music` to features list, usage parameters, feature parity table, and "Choosing Between Versions" section

**PR #49 - Add Music Directory Structure Example**
- Added Music directory structure example to README

**PR #50 - Add Missing Genre Directory Structure Examples**
- Added Tutorials, Fitness, and Surf directory structure examples to README

**PR #51 - Add Music Usage Example**
- Added "Rip a music disc" usage example to README Examples section

**PR #52 - Session Notes for PRs #46-#51**
- Added PRs #46-#51 to 2026-02-22 session notes in CLAUDE.md

**PR #53 - Fix Series Disc Subdirectory Removal**
- `Remove-Item` on empty Disc subdirectory failed with "in use" because PowerShell's working directory was still inside it
- Added `cd $seriesSeasonDir` before `Remove-Item` so the working directory moves to the parent season folder first
- Applied to both scripts

**PR #54 - Fix Composite Mega-File Detection**
- Old heuristic (`largest >= 2x second-largest`) failed when one episode was much longer than others
- New heuristic checks if the largest file is within 70-130% of the sum of all other files
- Since a composite is all episodes concatenated, its size should closely match the total
- Applied to both scripts

**PR #55 - Session Notes & Changelog for PRs #52-#54**
- Added PRs #52-#54 to CLAUDE.md session notes and CHANGELOG.md

**Cleanup:**
- Deleted 6 stale merged remote branches (feature/1-*, feature/2-*, etc.)
- Deleted stale local branch `fix/1-composite-megafile-detection`
- Deleted `feature/stroop-test-web-app` branch (PR #13 already closed, unrelated to ripdisc)

**Work In Progress:**
- None — all PRs merged, working tree clean, no stale branches

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation (see Feature Parity table in README)

---

### 2026-02-23 - Auto-Discover Disc Metadata

**Problem:**
`rip-disc.ps1` required `-title` as a mandatory parameter, plus manual flags like `-Series`, `-Season`, `-Bluray`. Unlike `rip-audio.ps1` which auto-detects metadata, video disc rips required the user to know and type all metadata upfront.

**Solution:**

**PR #57 - Add Auto-Discovery of Disc Metadata via MakeMKV + TMDb**

When `-title` is omitted, the script now auto-discovers disc metadata:

1. Reads disc info via `makemkvcon -r info` (disc name, type, title count)
2. Cleans the disc name (strips suffixes like `_D1`, `_WS`, `_DISC2`, replaces underscores, title-cases)
3. Searches TMDb (The Movie Database) API for the cleaned title
4. Auto-populates `-title`, `-Bluray`, `-Series`, `-Season`, `-Disc` from results
5. Prompts for confirmation (Accept / Edit / Abort)

When `-title` is provided, only Blu-ray format auto-detection runs (quick info query).

**Parameter changes:**
- `-title` changed from `[Parameter(Mandatory=$true)]` to `[Parameter()]` with default `""`

**New functions added to `rip-disc.ps1`:**

| Function | Purpose |
|----------|---------|
| `Get-DiscInfo` | Runs `makemkvcon -r info` and parses CINFO/TINFO fields (disc type, name, volume label, per-title duration/chapters/size) |
| `Clean-DiscName` | Strips suffixes, extracts season/disc hints via regex, replaces underscores, title-cases |
| `Search-TMDb` | Queries TMDb multi-search API (`search/multi`), filters to movie/tv, presents top 5 for user selection |

**Auto-detection matrix:**

| Parameter | Auto-detected? | Source |
|-----------|---------------|--------|
| `-title` | Yes | TMDb → cleaned disc name → manual fallback |
| `-Bluray` | Yes | MakeMKV CINFO:1 disc type |
| `-Series` | Yes | TMDb `media_type: "tv"` |
| `-Season` | Partial | Regex from disc name (e.g. `S01`, `Season 1`) |
| `-Disc` | Partial | Regex from disc name (e.g. `D2`, `Disc 2`) |
| Genre flags | No | Always manual (`-Documentary`, `-Music`, etc.) |

**Other changes:**
- `$makemkvconPath` moved from configuration section to immediately after param block (needed before discovery)
- `$discSource` built once in discovery section, reused by Step 1 MakeMKV rip
- Disc format shown in "Ready to Rip" confirmation display when discovered
- README updated: `-title` marked optional, new Auto-Discovery section with TMDb API key setup, feature parity table updated

**Files changed:**
- `rip-disc.ps1` — All discovery logic, parameter changes, new functions
- `README.md` — Documentation for auto-discovery feature and TMDb setup

**Files NOT changed:**
- `continue-rip.ps1` — Resumes from existing files, title always known, stays mandatory

**PR #60 - Fix Disc Discovery Parsing and Add Coffee Link**

Three fixes:

**1. Get-DiscInfo regex parsing fix:**
- Removed `$` anchors from all regex patterns — they fail on Windows `\r` line endings because MakeMKV outputs `\r\n` but PowerShell keeps the `\r` in captured strings
- Added `.Trim()` to each line before matching
- Added `ErrorRecord` filtering — `2>&1` merges stderr as `ErrorRecord` objects, not strings, which caused silent match failures

**2. Disc source default fix:**
- Changed default from `dev:$driveLetter` (`dev:D:`) to `disc:0` when no `-DriveIndex` specified
- `disc:0` lets MakeMKV auto-find the first available optical drive
- The old `dev:D:` assumed D: was always the optical drive, which isn't always correct

**3. Buy Me a Coffee link:**
- Added `https://buymeacoffee.com/stephenbeale` nudge to all successful completion outputs
- Three locations: normal completion (rip-disc.ps1), queue completion (rip-disc.ps1), continue completion (continue-rip.ps1)
- Styled: gray message text, cyan URL

**Files changed:**
- `rip-disc.ps1` — Regex fix, disc source fix, coffee link (2 locations)
- `continue-rip.ps1` — Coffee link (1 location)

**PR #62 - Show Drive Hint During Auto-Discovery**
- Display which drive is being scanned (e.g. "first available drive (disc:0)" or "G: ASUS external (disc:1)")
- Moved status message out of `Get-DiscInfo` into callers for context-appropriate messaging

**PR #63 - Silence Disc Format Check When -title Is Provided**
- Removed "Reading disc info..." message when `-title` is provided (confusing — looked like full discovery)
- Skip disc query entirely if both `-title` and `-Bluray` are provided (nothing to detect)

**PR #64 - Skip Slow Disc Query When -title Is Provided**
- Removed `Get-DiscInfo` call entirely from the title-provided path
- `makemkvcon info` takes 30-60+ seconds — not worth it just for Blu-ray auto-detection
- Users should pass `-Bluray` manually when providing `-title`
- Added "(This may take a minute while MakeMKV reads the disc)" hint for discovery mode

**PR #65 - Fix -Drive Parameter Ignored and Variable Collision**

Two bugs:

**1. `-Drive` parameter being ignored:**
- PR #60 changed default `$discSource` from `dev:$driveLetter` to `disc:0` for all cases
- This meant `-Drive G:` was ignored — MakeMKV always used disc:0 (D: internal)
- Fix: restored `dev:$driveLetter` as default for ripping; discovery mode (no title) uses separate `$discoverySource = "disc:0"` variable

**2. `$discType` variable name collision:**
- Local `$discType` ("Main Feature") in the "Ready to Rip" block collided with `$script:DiscType` (MakeMKV disc type like "DVD disc")
- PowerShell variables are case-insensitive, and at script scope `$discType` IS `$script:DiscType`
- Caused "Disc Format: Main Feature" instead of actual disc type
- Fix: renamed local variable to `$discTypeLabel`

**Files changed:**
- `rip-disc.ps1` — All fixes above

**PR #67 - Fix Blu-ray Subtitle Handling**
- Blu-ray PGS subtitles were getting burned in despite `--subtitle-burned=none`
- Replaced `--all-subtitles --subtitle-burned=none` with `--subtitle scan --subtitle-burned`
- Now scans for forced/foreign-language subs only (e.g. alien dialogue in English films) and burns those in
- Full PGS subtitle tracks are excluded entirely
- Removed the try-then-retry fallback pattern — no longer needed
- DVD behaviour unchanged (all subtitles as separate tracks)
- Updated in 3 places per script: main encoding, recovery script generation, comments

**Files changed:**
- `rip-disc.ps1` — Subtitle handling fix
- `continue-rip.ps1` — Subtitle handling fix

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-01 - Blu-ray Output Directory, Rename Fix & MakeMKV Progress

**PR #69 - Route Blu-ray Output to Dedicated Directory**
- `-Bluray` now routes to `<OutputDrive>:\Bluray\<title>\` instead of `DVDs\<title>\`
- All `$contentType` chains show "Blu-ray" as type label (Get-TitleSummary, Write-Log, Ready to Rip display)
- Queue entry includes `Bluray = [bool]$Bluray` for downstream processing
- Genre flags still take priority (e.g. `-Documentary -Bluray` goes to `Documentaries\`)
- Applied to both scripts

**PR #70 - Fix Double-Prefix in File Renaming**
- MakeMKV names files like `Southpaw_t01.mp4` (title + underscore)
- Prefix check only looked for `Title-*` (hyphen), missing `Title_*` (underscore)
- Result: `Southpaw_t01.mp4` got double-prefixed to `Southpaw-Southpaw_t01.mp4`
- Fixed all 6 prefix checks (3 per script) to also match `Title_*`

**PR #71 - Replace Underscore with Hyphen in Renamed Files**
- PR #70 skipped `Title_*` files entirely, leaving underscore separators in filenames
- Now detects `Title_*` and replaces the underscore with a hyphen: `Southpaw_t01.mp4` -> `Southpaw-t01.mp4`
- Applied to all 3 prefix paths (Disc 1, Extras, Disc 2+) in both scripts

**PR #72 - Stream MakeMKV Output to Console**
- MakeMKV rip output was silently captured — `$makemkvOutput = ... | Tee-Object` swallowed console output
- Removed variable assignment and piped through `ForEach-Object { Write-Host $_ }` for real-time streaming
- `Tee-Object -Variable makemkvFullOutput` still captures everything for error analysis

**Files changed:**
- `rip-disc.ps1` — Bluray directory routing, content type labels, queue entry, prefix rename logic, MakeMKV output streaming
- `continue-rip.ps1` — Bluray directory routing, content type labels, prefix rename logic

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed
- Test PR #75 drive fix on next real rip (disc:N mapping + drive wake-up)

---

### 2026-03-04 - Drive Path Fix & Documentation

**PR #73 - Update README, CHANGELOG, and Session Notes**
- Added Blu-ray directory structure example to README
- Updated `-bluray` usage description and feature parity table
- Fixed session date from 2026-02-24 to 2026-03-01
- Added PR #72 (MakeMKV streaming) to changelog and session notes

**PR #75 - Fix MakeMKV Drive Device Path**
- `Win32_CDROMDrive.DeviceID` is a PNP device ID (e.g. `USBSTOR\CDROM&VEN_...`), not `CdRomN`
- `dev:\\.\<PNP-DeviceID>` caused "Unknown device" / "Failed to open disc" for all USB drives
- Replaced with `disc:N` format using the drive's position in WMI `Win32_CDROMDrive` enumeration
- Added drive wake-up (`Test-Path`) before WMI query and before Step 1 to spin up dormant USB drives
- Error output now lists all available drives when target drive not found
- Untested on real disc — user's drives were all in use; to verify next session

**Files changed:**
- `rip-disc.ps1` — Drive source mapping rewrite, drive wake-up
- `README.md` — Blu-ray directory structure, feature descriptions
- `CHANGELOG.md` — 2026-03-01 and 2026-03-04 sections
- `CLAUDE.md` — Session notes

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Test PR #75 drive fix on next real rip
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-05 - Drive Index: WMI Retry then Direct MakeMKV Query

**Problem:**
PR #75 mapped drive letters to `disc:N` using the drive's position in WMI `Win32_CDROMDrive` enumeration. WMI ordering and MakeMKV's internal enumeration can disagree — particularly for USB drives — causing MakeMKV to rip from the wrong drive.

**PR #82 merged: `fix/makemkv-drive-index`**
- Initial fix: WMI-first approach with auto-retry fallback
- Added `Find-MakeMkvDriveIndex` helper that queries `disc:9999` to map drive letters to MakeMKV disc indices
- Added README explanation of why the Disc1 subdirectory is used

**PR #83 open: `fix/drive-index-direct`** (branch: `fix/drive-index-direct`, CodeRabbit passing)
- Replaces WMI-first-then-retry with direct `disc:9999` MakeMKV query every time
- Removes WMI lookup entirely, removes retry block, removes `Find-MakeMkvDriveIndex` helper
- Simpler, correct, and consistent — always asks MakeMKV which index owns a given drive letter
- Awaiting human approval to merge (branch protection)

**CodeRabbit comments on PR #83 (outside diff — address before or after merge):**
1. Harden `disc:9999` exit-code check: if `$LASTEXITCODE -ne 0` after the MakeMKV call, write error and `exit 1`
2. Normalize drive letter comparison: `TrimEnd('\')`, `TrimEnd(':')`, `ToUpperInvariant()` on both sides
3. Add `disc:0` fallback when `-Drive` was not explicitly passed (preserves single-drive backward compat)

**Files changed in PR #83:**
- `rip-disc.ps1` — Drive index logic replaced with direct `disc:9999` query

**Work In Progress:**
- PR #83 open — awaiting merge

**Outstanding Work for Future Sessions:**
- Merge PR #83 (`fix/drive-index-direct`) — CodeRabbit passing, needs human approval
- Consider applying CodeRabbit's 3 hardening suggestions to the merged code as a follow-up PR
- Test drive index fix on next real disc rip
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-11 - Config System, Setup Wizard, and v1.0.0 Release

**Problem:**
All tool paths (MakeMKV, HandBrake, TMDb API key, etc.) were hard-coded in both scripts. A new user downloading the repo would need to manually edit the scripts before they could run anything — not shareable.

**PR #91 merged: `feature/config-system`**

New files added:

| File | Purpose |
|------|---------|
| `setup.ps1` | Interactive first-run wizard — guides through Chocolatey install, manual download fallback, drive detection, TMDb API key setup; writes `ripdisc-config.json` |
| `Load-Config.ps1` | Shared config loader — reads `ripdisc-config.json`, falls back to auto-detection via registry, PATH scan, and common install paths |
| `ripdisc-config.sample.json` | Sample config committed to repo — users copy and fill in |
| `Start.bat` | Double-click entry point — runs `setup.ps1` with `-ExecutionPolicy Bypass` so users never need to touch PowerShell settings |

Changes to existing scripts:

- All hard-coded paths in `rip-disc.ps1` and `continue-rip.ps1` replaced with config variables loaded from `Load-Config.ps1`
- `ripdisc-config.json` added to `.gitignore` (user-specific, never committed)
- README updated with "Getting Started" section aimed at new users, linking to `Start.bat` and `setup.ps1`

End result: a friend can download the zip, double-click `Start.bat`, follow the wizard, and be ready to rip.

**GitHub Release v1.0.0 published**
- URL: https://github.com/stephenbeale/ripdisc/releases/tag/v1.0.0
- Includes a self-contained C# exe (65MB) as a downloadable release asset
- Release notes include getting started instructions

**PowerShell Gallery publishing**
- Considered and started; abandoned at user's request

**Work In Progress:**
- PR #83 open: `fix/drive-index-direct` — CodeRabbit passing; needs human approval (branch protection)

**Outstanding Work for Future Sessions:**
- Merge PR #83 (`fix/drive-index-direct`) — CodeRabbit passing, needs human approval
- Real-disc testing of drive index fix still pending (any rip will exercise this)
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-11 - Revert Drive Lookup to WMI (PR #92)

**Problem:**
PR #83 (`fix/drive-index-direct`) replaced WMI enumeration with a direct `disc:9999` MakeMKV query to map a drive letter to a `disc:N` index. The MakeMKV query approach had a critical flaw: it physically scanned all optical drives in the system, including drives that were mid-rip in concurrent sessions. This caused interference between concurrent rips.

**PR #92 merged: `fix/revert-drive-lookup`**
- Reverted the drive letter to `disc:N` mapping back to WMI `Win32_CDROMDrive` enumeration
- WMI queries Windows only — it never touches physical drives or disturbs active rip sessions
- The `disc:9999` MakeMKV approach (from PR #83) is fully removed

**Current approach (post PR #92):**
- Drive letter is mapped to `disc:N` index by enumerating `Win32_CDROMDrive` via WMI
- WMI drive order is assumed to match MakeMKV's internal `disc:N` enumeration for the same hardware

**Implication for PR #83:**
- PR #83 (`fix/drive-index-direct`) is now OBSOLETE — the approach it implemented has been reverted
- PR #83 should be closed on GitHub as superseded by PR #92

**Work In Progress:**
- None — main branch clean, fully synchronized with origin/main

**Outstanding Work for Future Sessions:**
- Close PR #83 on GitHub (obsolete — superseded by PR #92 revert to WMI)
- Real-disc testing to verify WMI index mapping works correctly for G: (USB DVD drive)
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-23 - Drive Index Mapping Fix (PRs #96–#98)

**Problem:**
WMI `Win32_CDROMDrive` enumeration order does not reliably match MakeMKV's internal `disc:N` ordering, particularly for USB optical drives. This caused the script to rip from the wrong physical drive when using `-Drive G:`.

**PR #96 merged: `fix/drive-index-makemkv-query`**
- Replaced WMI `Win32_CDROMDrive` drive lookup with a direct MakeMKV `disc:9999` query
- `disc:9999` causes MakeMKV to enumerate all connected optical drives and return their drive letters alongside their `disc:N` indices
- Script parses the output and maps the requested drive letter to the correct `disc:N` index
- WMI is removed entirely from the drive lookup path

**PR #97 merged: `fix/drive-enumeration-display`**
- At startup, all drives detected by MakeMKV are listed (e.g. `disc:0 D: HL-DT-ST`, `disc:1 G: GP75N`)
- An arrow marker (`-->`) highlights the selected drive
- Allows the user to visually verify the correct physical drive is being used before the rip begins

**PR #98 merged: `fix/skip-busy-drives`**
- Changed from `disc:9999` (probes all drives simultaneously) to iterating `disc:0`, `disc:1`, etc. individually
- Detects running `makemkvcon` processes and reads their command lines to identify which `disc:N` indices are in use
- Skips busy indices to avoid interfering with concurrent rip sessions
- Stops iterating once the target drive letter is found — no unnecessary probing

**Verified working:**
- Tested with `-Drive G:` — correctly mapped to `disc:1` (GP75N)
- Rip completed successfully with 18 titles

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Investigate CSS authentication issue: "Muriel's Wedding" on G: (USB DVD) fails with CSS SCSI errors via makemkvcon CLI but rips fine in MakeMKV GUI — run `makemkvcon mkv disc:2 all "C:\Video\test" --minlength=120` directly to isolate CLI vs GUI difference
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-03-23 (continued) - Stuck Sector Detection and Bug Fixes (PRs #99–#101)

**PR #99 merged: `fix/detect-stuck-sectors`**
- Monitors makemkvcon stdout for repeated error messages at the same byte offset
- Kills makemkvcon process when 5 consecutive errors at the same offset are detected
- Prevents indefinite stalls caused by physically damaged discs (e.g. bad sectors causing hardware timeouts)
- Tested scenario: Shutter Island disc was triggering repeated hardware timeout errors without any progress

**PR #100 merged: `fix/stuck-sector-ps51-compat`**
- Original implementation used `System.Threading.Thread` and `System.Collections.Concurrent.ConcurrentQueue`
- PS 5.1 does not support these .NET threading primitives in the way the code used them
- Rewrote using a `cmd.exe` wrapper around makemkvcon with stdout-only redirect and synchronous line reading

**PR #101 merged: `fix/busy-drive-and-exitcode`**
Two bugs fixed in one PR:
1. `-not $isBusy` guard removed — this condition blocked selecting a drive that had previously been "busy" in the process scan, even though the drive was now free and available for ripping
2. `cmd.exe` wrapper replaced with direct `Process` object execution — the cmd.exe approach returned exit code 2 due to path quoting issues; direct Process invocation passes arguments correctly and respects the actual makemkvcon exit code

**Current Architecture Notes:**
- MakeMKV execution: direct `System.Diagnostics.Process` object, stdout-only redirect, synchronous `ReadLine()` loop
- Stuck sector detection: 5 consecutive identical offset errors triggers process kill; counter resets on any new offset
- Drive lookup: `disc:0`, `disc:1`, etc. iterated individually; busy indices (from live `makemkvcon` process cmdlines) are skipped; stops on first match
- Startup display: all detected drives listed, selected drive marked with `-->`

**Work In Progress / Needs Testing:**
- Stuck sector detection (PRs #99–#101) needs real-world validation on a damaged disc
- Shutter Island disc was the test case — hardware timeout errors were occurring; retry the rip to confirm the kill logic fires correctly
- Direct Process execution (PR #101) replaces the cmd.exe wrapper — verify exit code is captured correctly on both success and failure

**Outstanding Work for Future Sessions:**
- Validate stuck sector detection on a genuinely damaged disc
- Investigate CSS authentication issue: "Muriel's Wedding" on G: (USB DVD) fails with CSS SCSI errors via makemkvcon CLI but rips fine in MakeMKV GUI
- Port missing features to C# implementation (see Feature Parity table in README)

---

### 2026-04-15 - Continue From Existing Files Option (PR #104)

**Problem:**
When the MakeMKV output directory already contained files from a previous rip attempt, users were presented with only two options: delete the directory and start over, or use a suffixed directory. There was no way to skip the MakeMKV rip entirely and resume directly at Step 2 (HandBrake encoding) using the files that were already there.

**Solution:**

**PR #104 merged: `feature/continue-from-existing`**

Added a third option to the "MakeMKV output directory already exists" prompt in `rip-disc.ps1`:

- Option [1] — Delete and re-use directory (existing behaviour, unchanged)
- Option [2] — Use a suffixed directory, e.g. `title-1` (existing behaviour, unchanged)
- Option [3] — Continue with existing files (skip MakeMKV rip) — NEW

When the user selects option [3]:
- The MakeMKV rip is skipped entirely
- Disc eject is skipped (no disc was ripped, so there is nothing to eject)
- Existing `.mkv` files in the output directory are picked up as the source for Step 2
- HandBrake encoding and all subsequent steps (organize, rename, cleanup) proceed normally

**Use case:**
MakeMKV completed successfully in a previous session but the script was interrupted before HandBrake finished. The user can restart `rip-disc.ps1` and select option [3] to resume encoding from the already-ripped files, avoiding a full re-rip.

**Files changed:**
- `rip-disc.ps1` — Added option [3] branch to the existing-directory prompt; skips `Invoke-MakeMKV` and disc eject when selected; continues to HandBrake with files from the existing directory

**Work In Progress:**
- None — PR merged, working tree clean, main fully synchronised with origin/main

**Outstanding Work for Future Sessions:**
- Validate stuck sector detection on a genuinely damaged disc
- Investigate CSS authentication issue: "Muriel's Wedding" on G: (USB DVD) fails with CSS SCSI errors via makemkvcon CLI but rips fine in MakeMKV GUI
- Port missing features to C# implementation (see Feature Parity table in README)
- Consider whether option [3] should also be added to `continue-rip.ps1` for symmetry

---

### 2026-08-10 - Stuck Sector Watchdog False Positive on CSS Auth Errors (PR #105)

**Problem:**
Two rips failed at Step 1 with `No MKV files were created`:
- "Who Framed Roger Rabbit" on G:
- "The Great Gatsby 2013" on H:

Both were preceded by `SCRAMBLED SECTOR WITHOUT AUTHENTICATION` SCSI errors, at offsets 1048576 and 589824 respectively.

**Root cause:**
Drive H: (`DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933`) cannot CSS-authenticate its disc. `makemkvcon` enumerates every optical drive at startup, so even a rip told to use `disc:1` (drive G:) still hit H: and emitted repeated errors like:

```
Error 'Scsi error - ILLEGAL REQUEST:READ OF SCRAMBLED SECTOR WITHOUT AUTHENTICATION' occurred while reading 'DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933' at offset '589824'
```

The stuck-sector watchdog counted those startup-enumeration errors — which came from a *different* drive — and killed `makemkvcon` roughly 2 seconds in, before the real rip had started. `$makemkvExitCode = if ($wasKilledForStuck) { 0 }` then masked the kill as success, so the CSS-specific error branch never fired and the user only saw the generic "No MKV files were created".

**Solution:**

**PR #105: `fix/stuck-detector-css-false-positive`**

Five changes, all in `rip-disc.ps1`:

1. **Target drive name captured** — new `$script:TargetDriveName`, set from `$matchedDrv.Name` when the drive list resolves, so read errors can be attributed to a specific drive. Empty on the `-DriveIndex` path (the drive list is never seen there).
2. **Drive list cache skipped on enumeration errors** — `makemkv-drive-cache.txt` was storing raw `info disc:9999` output including `MSG:2003` lines. A drive that errors during enumeration may be missing or wrong in the `DRV:` list, and that bad mapping was reused for the full 5-minute TTL. The cache is now written only when the enumeration is error-free.
3. **Watchdog gated on rip-started plus drive attribution** — offset errors count toward the stuck counter only once `$ripStarted` is true (set on `Saving N titles`, `Current progress`, `Current operation`, or `Title #`) *and* the error names our drive, parsed from `occurred while reading '<drive>' at offset '<n>'`. Errors naming another drive are still printed but never counted. A separate pre-rip escape hatch (50 consecutive errors, `$wasKilledForAuth`) stops an unauthenticatable disc hanging forever at enumeration. Threshold stays 5; non-error lines still reset the counter.
4. **Exit code judged by salvaged files** — no longer forced to 0. The output directory is checked for `*.mkv` right after the process exits; exit 0 only when a stuck kill actually salvaged files, otherwise an explicit non-zero code (`Kill()` leaves `$proc.ExitCode` unhelpful) so error analysis runs.
5. **CSS detection un-nested** — a new branch matching `SCRAMBLED SECTOR WITHOUT AUTHENTICATION` anywhere in the output, evaluated *before* the `Failed to open disc` check (which became `elseif`). It names the offending drive when parseable and warns that it may be a different drive from the one being ripped, suggesting the user remove discs from other optical drives and retry.

**Relationship to the long-standing "Muriel's Wedding" item:**
This is the same signature as the open item carried in these notes since 2026-03-23 — *"Muriel's Wedding on G: fails with CSS SCSI errors via makemkvcon CLI but rips fine in the MakeMKV GUI"*. That difference is now at least partly explained: the GUI has no watchdog killing it, so it survives the same enumeration errors that were aborting the CLI rip. Whether the underlying drive-level CSS authentication failure on H:/G: is itself fixed is **still unverified** — this change stops the script from killing the rip, it does not make an unauthenticatable drive authenticate.

**Files changed:**
- `rip-disc.ps1` — All five changes above
- `CHANGELOG.md` — 2026-08-10 section
- `CLAUDE.md` — These session notes

**Files deliberately NOT changed:**
- `continue-rip.ps1` — never runs MakeMKV, so none of this applies
- C# implementation — not ported

**Testing status — IMPORTANT:**
Parse-checked only (`[System.Management.Automation.PSParser]::Tokenize` reports 0 errors), and the UTF-8 BOM on `rip-disc.ps1` was verified intact after editing. **This fix has NOT been runtime-tested** — no disc was mounted in any drive during the session.

Verification step for the next session: remove the disc from H:, re-run the G: rip, and confirm it completes.

**Work In Progress:**
- None — PR #105 merged, working tree clean

**Outstanding Work for Future Sessions:**
- Verify the CSS fix on a real rip (remove the disc from H:, re-run the G: rip, confirm completion)
- Investigate why H: (`GP75N 1.01 K0MMB391933`) cannot authenticate — check region code and RPC setting on the drive
- Validate stuck sector detection on a genuinely damaged disc
- Port missing features to C# implementation (see Feature Parity table in README)
- Auto-discovery is PowerShell only — add to C# if needed

---

### 2026-08-10 (continued) - Eject False Timeout (PR #106)

**Problem:**
Rips were reporting `Disc eject timed out after 2 attempts - please eject manually` while the disc had, in fact, already been ejected — the tray was open. It happened on many discs across all three drives.

**Root cause — the timeout was measuring the wrong thing:**
The eject ran on a background job:

```powershell
$ejectJob = Start-Job -ScriptBlock { ... InvokeVerb("Eject") } -ArgumentList $driveLetter
$ejectCompleted = $ejectJob | Wait-Job -Timeout 15
```

`Start-Job` spawns an entire child `powershell.exe`. Measured on this machine while 12 concurrent `HandBrakeCLI` encodes had the CPU pinned at 100%, a `Start-Job` with a trivial `{ 1 }` body took **18.0s, 25.7s and 33.2s** across three runs. The 15s deadline was therefore consumed by process startup, before the eject had any chance to report back. The child eventually ran and the disc came out; the script had already declared failure.

The saturation is self-inflicted: this is the same tool running several rips at once, so the more concurrent sessions, the worse it gets.

**Evidence in `C:\Video\logs`:**
- All of June 2026 — zero eject timeouts, across D:, E: and G:
- 2026-08-10 — every rip up to 15:14 succeeded; 8 of the 10 after it failed, as concurrent encodes accumulated through the afternoon
- Failures hit E:, G: and H: equally — not drive-specific and not disc-specific, which rules out drive firmware and the CSS problems tracked above
- Runaway Jury (15:16) succeeded on attempt 2 — a marginal case straddling the deadline
- Time from "Retrying eject (attempt 2)" to the failure log was 23–33s against a 15s timeout, the extra being `Stop-Job`/`Remove-Job` on a job whose child was still starting

**Solution (`rip-disc.ps1` only):**

1. **Eject issued in-process as a direct device IOCTL.** `IOCTL_STORAGE_EJECT_MEDIA` (`0x2D4808`) via a small P/Invoke class `RipDiscEject`, compiled lazily and guarded by `if (-not ('RipDiscEject' -as [type]))`. No child process. It also only touches the target drive — the old `Shell.Application` verb made Explorer re-enumerate every optical drive, which stalls when a sibling drive is mid-rip in a concurrent session.
2. **Success confirmed by observation, not by a deadline.** `System.IO.DriveInfo.IsReady` is polled (1s interval, 20s ceiling) until the drive goes not-ready. Measured at 3–24ms per check even at 100% CPU, so polling is effectively free. `Win32_CDROMDrive.MediaLoaded` was rejected as the signal — it took **13 seconds** on E: under the same load.
3. **Failure message reworded** — a genuine failure is no longer a timeout, so the dialog now says "Could not eject the disc … Please eject it manually."

**Verified — runtime-tested, unlike PR #105:**
- BOM intact (`EF BB BF`), `PSParser::Tokenize` reports 0 errors
- The embedded C# was extracted from the saved file and compiled — signature `Boolean Eject(String driveRoot, Int32& error)`
- Live eject on H: **at 100% CPU load**: IOCTL returned `ok=True` in 1309ms, eject confirmed after 267ms with **zero poll iterations**. The old code would have called this a timeout.
- Drive-root normalisation checked for `H:`, `H` and `H:\` — all produce `\\.\H:`

**Scope — checked, deliberately not changed:**
- `continue-rip.ps1` — has no eject logic at all
- `rip-audio.ps1` (separate `ripaudio` repo) — ejects via `Shell.Application` with no timeout and no confirmation, so it cannot produce this false negative
- A sweep for `Start-Job|Wait-Job|Start-Sleep|-Timeout|WaitForExit` across all three scripts found **this was the only `Start-Job`/`Wait-Job` in the codebase**. Everything else that waits either polls real state or counts discrete events (the stuck-sector watchdog counts consecutive identical-offset errors; file-lock retries are attempt-count-based) rather than trusting a wall clock against unpredictable process-spawn cost. This was a one-site bug, not a systemic pattern.

**Files changed:**
- `rip-disc.ps1` — Eject rewrite
- `CHANGELOG.md`, `CLAUDE.md`

**Work In Progress:**
- None

**Outstanding Work for Future Sessions:**
- Unchanged from the list above — none of those items interact with this fix
- Consider whether the C# port items, carried since 2026-02-16 with no commit touching `RipDisc/*.cs` since, should be formally abandoned rather than re-listed every session

---

### 2026-08-11 - Eject Fix Confirmed in Production; continue-rip.ps1 Usability (PR #107)

**Eject fix (PR #106) — verified on real rips, condition closed:**

Three concurrent rips were started at 06:13–06:14 (`disc:0`, `disc:1`, `disc:2`), which is the CPU-saturation condition that caused the original false timeouts. Both ejects that ran succeeded:

| Drive | MakeMKV complete | Ejected | Elapsed |
|-------|------------------|---------|---------|
| G: | 06:18:54 | 06:19:00 | 6s |
| H: | 06:24:23 | 06:24:29 | 6s |

No `Retrying disc eject` lines and no timeout warnings, against 8 failures out of 10 rips under the same concurrency the previous afternoon. The 6 seconds matches the measured breakdown: one-time `Add-Type` compile (~6s under load) plus a 1.3s IOCTL plus immediate confirmation. **This item is now closed** — it no longer needs carrying forward.

Worth noting for future work: the ~6s is almost entirely the lazy `Add-Type` compile, not the eject. If that ever matters, compiling `RipDiscEject` at script start would move the cost off the eject path — but there is no deadline any more, so it does not currently matter.

**PR #107 — `continue-rip.ps1` usability and a Step 3 directory bug**

Authored across two sessions; 525 insertions / 85 deletions.

*The real bug — Step 3 renaming files in the wrong directory:*
Step 3 ran a bare, unchecked `cd $finalOutputDir`. The entire organize block renames and moves files relative to the *current* directory, so when that `cd` failed the script carried on regardless and renamed whatever was in the directory it had been launched from. This actually hit the repo working directory in practice. Now:

```powershell
Set-Location -LiteralPath $finalOutputDir -ErrorAction Stop   # in a try/catch
# then a post-check comparing (Get-Location).Path to the target before renaming
```

*Usability changes:*
- `title` and `FromStep` optional — omitting them shows a step menu instead of failing on a mandatory-parameter prompt
- `-Drive` / `-DriveIndex` accepted and ignored, so a failed `rip-disc.ps1` command line pastes straight into `continue-rip.ps1` without editing
- Every step always listed with its prerequisites; the chosen step can be switched at the prompt
- Prerequisite checks extracted to `Test-StepPrerequisites`; `Stop-Prerequisite` became `Write-PrerequisiteFailure`, returning `$false` rather than `exit 1`, so a failed check re-offers the step list instead of killing the session
- `-Yes` still exits 1 on a failed prerequisite, so non-interactive behaviour is unchanged

**Process note — two sessions on one branch:**
Two Claude sessions worked `feature/continue-rip-usability` simultaneously. One session's `git add` was swept up by the other's commit (`bb09335`), and a scratch file `continue-rip-head.ps1` appeared and vanished mid-workflow. Nothing was lost, but the near-miss is worth remembering: **one branch, one session.** A merge with `--delete-branch` during that window would have truncated the other session's work.

**Also worth knowing:** `gh pr review --approve` can never succeed on these PRs — GitHub rejects self-approval on your own PR ("Can not approve your own pull request"). The `CLAUDE.md` git workflow above lists approval as step 5; in practice it is always skipped, and merges land with zero approving reviews. Recording an approval would need a second account or a bot reviewer.

**Files changed:**
- `continue-rip.ps1` — Step 3 guard and usability rework
- `CHANGELOG.md`, `CLAUDE.md`

**Work In Progress:**
- None

**Outstanding Work for Future Sessions:**
- Verify the CSS fix (PR #105) on a real rip — still parse-checked only, never runtime-tested
- Investigate why H: (`GP75N 1.01 K0MMB391933`) cannot authenticate — check region code and RPC setting
- Validate stuck sector detection on a genuinely damaged disc
- Decide the fate of the C# port items, carried since 2026-02-16 with no commit touching `RipDisc/*.cs` since

---

### 2026-08-11 (continued) - Step 3 Guard for rip-disc.ps1; Skip Already-Encoded Files (PRs #109, #110)

**Incident — every file in the repo renamed with a leading `-`:**
At ~06:13, `.gitignore`, `CHANGELOG.md`, `CLAUDE.md`, `continue-rip.ps1`, `LICENSE` and `Load-Config.ps1` were renamed to `-<filename>` — alphabetically, stopping at the `nul` file. Root cause: `continue-rip.ps1` Step 3's prefix-rename operates on the CURRENT directory, and `cd $finalOutputDir` failed silently, leaving the working directory wherever the script had been launched from — the repo — with an empty prefix. Files were restored; nothing was lost. This is exactly what the Step 3 guard in PR #107 fixes, and PR #109 below ships the same guard for `rip-disc.ps1`, which has the identical bare `cd`.

**Incident — a dry run started a real encode:**
While dry-running the user's command for testing, a piped blank line was read as pressing Enter at the confirmation prompt, and `continue-rip.ps1` began encoding for real. It was killed ~3 minutes in, after confirming by PID that the user's concurrent Disc 1 encode was left untouched. It truncated `F:\Documentaries\Metal A Headbanger's Journey-Disc 2\B2_t00.mp4` from 307 MB to 150 MB; the truncated file was deleted, and a later run by the user re-encoded it to 308.8 MB — fully repaired. The confirm prompt now uses `Read-Answer` (added in #107), which exits cleanly on an unanswered/EOF prompt rather than defaulting to "continue".

**PR #109 — `rip-disc.ps1` Step 3 working-directory guard**

Same class of bug as the `continue-rip.ps1` fix in #107, applied to `rip-disc.ps1`'s Step 3 (`STEP 3/4: Organizing files`), which has the identical bare, unchecked `cd $finalOutputDir`:

```powershell
try {
    Set-Location -LiteralPath $finalOutputDir -ErrorAction Stop
} catch {
    Stop-WithError -Step "STEP 3/4: Organize files" -Message "Cannot change directory to $finalOutputDir - $($_.Exception.Message)"
}
$currentPath = (Get-Location).Path.TrimEnd('\')
if ($currentPath -ne $finalOutputDir.TrimEnd('\')) {
    Stop-WithError -Step "STEP 3/4: Organize files" -Message "Refusing to organize: expected to be in $finalOutputDir but the working directory is $currentPath"
}
```

`-LiteralPath` also stops glob metacharacters in a disc title (`[`, `]`, `*`) from sending `Set-Location` to the wrong place. Preventative — this guard has not fired against a real failure, only verified by logic and comparison behaviour.

**PR #110 — `continue-rip.ps1` skips already-encoded files in Step 2**

Resuming a rip usually means encoding stopped part way through; re-encoding files that already finished wasted hours. Step 2 now partitions the MKV list and only encodes files with no matching MP4 in the output folder:
- New `-Force` switch (alias `-ReEncode`) re-encodes everything, including files that already have an MP4
- Each skipped file is printed and logged with the existing MP4's size; the message states this assumes the existing MP4s are complete
- "Nothing to encode" is reported and the script continues to Step 3 when every MKV already has an MP4
- Recovery script is generated only when there is something to encode, from the filtered list; `$recoveryScriptPath` starts `$null` and its post-encode deletion is guarded so `Test-Path` is never handed a null path
- The encode loop and its "file N of M" counters (console and log) are driven from the filtered list
- The pre-flight prerequisite summary reports which files will be skipped, or overwritten under `-Force`, plus a count of files to encode this run

Not yet exercised against a real HandBrake encode — verified against fixture files only. The skip trusts that an existing MP4 is complete: it checks existence, not integrity. A truncated MP4 from an interrupted encode (see the dry-run incident above) would be skipped as done. Possible hardening: compare duration via ffprobe, or flag outputs implausibly small relative to their source MKV.

**Process note — self-approval:**
GitHub refused self-approval on #107, #109 and #110 alike (`Review Can not approve your own pull request`); all three merged with no formal approval recorded. CodeRabbit passed on each. A concurrent session was also writing to this repo during the merges.

**Files changed:**
- `rip-disc.ps1` — Step 3 working-directory guard (#109)
- `continue-rip.ps1` — skip-already-encoded logic in Step 2 (#110)
- `CHANGELOG.md`, `CLAUDE.md`

**Work In Progress:**
- None

**Outstanding Work for Future Sessions:**
- Neither working-directory guard (rip-disc.ps1 or continue-rip.ps1) has fired against a real failure — logic and comparison behaviour verified only
- The skip-already-encoded feature has not been exercised against a real HandBrake encode — verified against fixture files only; consider hardening via an ffprobe duration check or a plausible-size check against the source MKV
- `F:\Documentaries\Metal A Headbanger's Journey-Disc 2\B7_t15.m4a` is 0 bytes with an audio-only extension — that title did not encode properly
- All 16 MKVs remain in `C:\Video\Metal A Headbanger's Journey-Disc 2\Disc1` (temp cleanup only runs after a clean end-to-end pass)
- Verify the CSS fix (PR #105) on a real rip — still parse-checked only, never runtime-tested
- Investigate why H: (`GP75N 1.01 K0MMB391933`) cannot authenticate — check region code and RPC setting
- Validate stuck sector detection on a genuinely damaged disc
- Decide the fate of the C# port items, carried since 2026-02-16 with no commit touching `RipDisc/*.cs` since

---

### 2026-08-17 - Blu-ray Burned-In Subtitle Incident; Guard Shipped (PRs #112, #113)

**Incident — Ocean Wonderland 3D ripped with burned-in PGS subtitles, unrecoverable:**

The user ran `rip-disc.ps1` for "Jean-Michel Cousteau Presents Ocean Wonderland 3D" with `-Documentary` but without `-Bluray`. Result: 3 MP4s (5.28 GB total) landed in `F:\Documentaries\Jean-Michel Cousteau Presents Ocean Wonderland 3D` with PGS subtitles burned into the picture.

Root cause: without `-Bluray` the DVD subtitle branch runs (`--all-subtitles --subtitle-burned=none`), and Blu-ray PGS tracks get burned in anyway despite that flag — which is the entire reason the `-Bluray` branch exists (PR #67; `--subtitle scan --subtitle-burned`).

Confirmed unrecoverable: burned-in subtitles are baked into the pixels, and `rip-disc.ps1` deletes the source MKVs via `Remove-Item -Recurse -Force`, which bypasses the Recycle Bin. Verified no MKVs survive under `C:\Video` for this title and nothing relevant is in the Recycle Bin. The only fix is a re-rip (~80 min: 26 min MakeMKV + 54 min HandBrake, per the original log timings). Log for reference: `C:\Video\logs\Jean-Michel Cousteau Presents Ocean Wonderland 3D_disc1_20260817_122444.log`.

Incidental finding worth checking on the re-rip: `_t00` and `_t02` were both 9.18 GB — near-certainly the 2D and 3D versions of the same feature. The script picked `_t02` as Feature and pushed `_t00` to `extras\`.

**PR #112 — quote the explorer.exe path**

`rip-disc.ps1` and `continue-rip.ps1` both called `Start-Process explorer.exe -ArgumentList $directoryToOpen` with the path unquoted. `Start-Process` joins `-ArgumentList` on spaces without quoting the elements, so a path like `C:\Video\Who Framed Roger Rabbit\Disc1` reached explorer.exe as three separate arguments and only the first token was treated as a path. Titles with spaces are the norm here, so this fired on most rips — impact was limited to the final "open the folder" convenience step, not the rip itself. Fixed by wrapping the path in embedded quotes, with `TrimEnd('\')` so a trailing backslash cannot escape the closing quote. Same defect class as a same-day fix in the `ripaudio` repo, where it was more severe (it broke the automatic handoff to `search-metadata.ps1`).

**PR #113 — Blu-ray mode guard**

New "BLU-RAY MODE GUARD" in both `rip-disc.ps1` and `continue-rip.ps1`. 139 insertions, 0 deletions (purely additive).

Design points:
- Triggers when any SINGLE ripped MKV exceeds 8.5 GB (DVD-9 capacity) and `-Bluray` was not passed. A title cannot exceed the disc it came from, so this is unambiguous.
- Deliberately per-file, NOT the sum. The project-manager agent originally proposed summing the MKVs; that would false-positive on ordinary DVDs because MakeMKV routinely emits the same feature as several titles — this very rip proves it (two 9.18 GB copies of the same feature). Unit test confirms 3×4 GB (12 GB total) does not trigger.
- In `rip-disc.ps1` it sits before the queue block, so `-queue` jobs carry the corrected flag through to the C# encoder.
- Offers: enable Blu-ray handling (default), continue in DVD mode, or abort. Abort keeps the MKVs and prints the `continue-rip.ps1 -FromStep 2 -Bluray` resume command.
- The default — including an unanswered/piped-EOF prompt — is the corrective, non-destructive option, deliberately chosen given the PR #109/#110 incident where a piped blank line was read as consent and started a real encode.
- `continue-rip.ps1` honours `-Yes` by auto-enabling rather than prompting.
- Output directory routing deliberately NOT changed: `$finalOutputDir` is resolved before Step 1 and already exists by the time the guard runs, so only subtitle handling flips. This is logged and shown on screen.

**Testing status:** `PSParser::Tokenize` reports 0 errors on both scripts; UTF-8 BOM confirmed intact on both (PS 5.1 requirement); 8/8 logic unit tests pass, including the per-file-vs-total false-positive case and the 8.5/8.6 GB boundary. **Not exercised against a real end-to-end rip.**

**Process note:** No CI is configured on this repo. CodeRabbit posted "review available on request" rather than reviewing (manual-trigger for repos under 10 stars) — it opted out, was never in-flight. Nothing force-merged. Per the standing self-approval limitation (GitHub blocks approving your own PR), `gh pr review --approve` was not attempted; an explanatory comment was left instead and the PR was merged directly.

**Files changed:**
- `rip-disc.ps1` — explorer.exe path quoting (#112), Blu-ray mode guard (#113)
- `continue-rip.ps1` — explorer.exe path quoting (#112), Blu-ray mode guard (#113)
- `CHANGELOG.md` — 2026-08-17 section (both PRs)

**Work In Progress:**
- None — both PRs merged, working tree clean, main synchronised with origin/main

**Action item for the user — Ocean Wonderland 3D re-rip not yet started:**
```
.\rip-disc.ps1 -title "Jean-Michel Cousteau Presents Ocean Wonderland 3D" -Bluray -Documentary -Drive E: -OutputDrive F
```
`F:\Documentaries\Jean-Michel Cousteau Presents Ocean Wonderland 3D` must be cleared first, or the prefix-rename and Feature-selection logic will collide with the existing burned-in files still sitting there. Caveat: `-Bluray` mode still burns in FORCED subtitles by design (`--subtitle scan --subtitle-burned`) — if this disc flags its subtitles as forced, the burn-in could recur. If that happens, the fix is a HandBrake-only re-encode with subtitles off (no re-rip needed), provided the MKVs are kept this time rather than deleted.

**Outstanding Work for Future Sessions:**
- **Five preventative fixes now shipped but never fired against a real failure** — both working-directory guards (#107/#109), the CSS stuck-sector false-positive fix (#105), stuck-sector detection itself (#99–#101), and now the Blu-ray guard (#113). Two of these (the Blu-ray guard and the skip-already-encoded feature, #110) will self-verify through ordinary use — no special test needed, just watch the next few rips. The CSS fix (#105) needs a deliberate test: remove the disc from H:, re-run a G: rip. The working-directory guards and stuck-sector detection will only prove themselves if the failure conditions recur naturally.
- Ocean Wonderland 3D re-rip — see action item above
- Verify the CSS fix (PR #105) on a real rip — still parse-checked only, never runtime-tested
- Investigate why H: (`GP75N 1.01 K0MMB391933`) cannot authenticate — check region code and RPC setting
- Validate stuck sector detection on a genuinely damaged disc
- The C# port items have now been carried since 2026-02-16 (over six months, 20+ sessions) with no commit touching `RipDisc/*.cs` in that entire span. The 2026-08-10 notes already suggested formally abandoning this rather than re-listing it every session; that suggestion is repeated here and should probably be acted on — either close it out explicitly as "PowerShell is the maintained implementation" or open a single tracking issue and stop carrying the bullet list in session notes
- **Cleanup candidate, not a git concern:** a stray `nul` file sits at the repo root and another at `.claude\nul` — both are the Windows reserved-device-name artifact (likely from an old `> nul` redirect that PowerShell wrote literally instead of discarding). Both are untracked and already covered by `.gitignore`, so they don't affect repo cleanliness from git's perspective, but they're real files on disk. First noted in the 2026-08-11 incident notes above (the alphabetical prefix-rename stopped at `nul`); not created this session. Low priority — reserved device names can't be addressed with an ordinary path; remove with the `\\?\` extended-path prefix, e.g. `Remove-Item '\\?\C:\Users\sjbeale\source\repos\ripdisc\nul'` and the equivalent for `.claude\nul`, whenever convenient

---

### 2026-08-24 - Documentary / Genre Series Mode for Multi-Disc Box Sets

**Problem:**
User has a physical boxset in hand: *Martin Scorsese Presents the Blues* — 5 discs, 7 feature-length episodes by different directors, one-or-two per disc. All 5 discs report the same or a near-identical disc label, so there's no way to distinguish them automatically — and `-Documentary` alone routes every disc to the same flat `Documentaries\<title>\` folder with Feature/extras logic that assumes a single film per disc. The largest-file-on-the-disc becomes "the Feature" and everything else gets shoved into `extras\` — wrong for a disc holding two films of near-equal length.

**Design:**
`-Series` already has everything a multi-disc box set needs — per-disc temp/output isolation (`Disc$Disc` subfolders, added in PR #24), composite mega-file detection (PR #31/#54), and a `-StartEpisode` parameter that existed but was never wired to anything. Rather than build a parallel system, combining `-Series` with a genre flag (`-Documentary`, `-Tutorial`, `-Fitness`, `-Music`, `-Surf`) now activates "genre series" mode, computed once as `$script:IsGenreSeries` right after config load and threaded through every place that already branches on `$Documentary`/`$Series` (title summary, window title, temp dir, final output dir, Ready-to-rip display, logging, Step 3 organize). Plain `-Series` (fiction TV, no genre flag) and plain `-Documentary` (single film) are both completely untouched — the new branch only fires when both are set.

**Disc-number ambiguity:** unchanged — the user already supplies `-Disc N` manually for every multi-disc rip (genre flags were never auto-detected from disc label in the first place; see the Auto-Discovery matrix above), so identical disc labels were never actually a blocker. Nothing new was needed here.

**Episode identity (the actual new logic) — Step 3 in both scripts:**
Every MKV on a genre-series disc is treated as an episode of equal standing (no Feature/extras split, matching how plain `-Series` already treats episodes). Files are renamed `<title>-E##.mp4` (or `<title>-S##E##.mp4` if `-Season` is given) and **moved up** out of the per-disc `Disc$Disc` subfolder into the shared title/season folder — unlike plain `-Series` mode, which deliberately leaves episodes nested in `Disc$Disc` forever (see the 2026-02-16 notes above; that's existing, intentional, unchanged behaviour for fiction TV). The now-empty `Disc$Disc` folder is removed (`cd` to the parent first, same fix as PR #53, or `Remove-Item` fails "in use").

**Resuming across sessions:** `-StartEpisode` still works as an explicit override, but is no longer required. When omitted, the script scans the shared target folder for the highest existing `-E##`/`S##E##` file and continues from `max + 1` — rip disc 1 today, disc 4 next week, no need to remember or compute where numbering left off. `Get-UniqueFilePath` (already used for extras-collision handling) is reused as a safety net in case auto-detection and reality disagree.

**Bug caught by the added unit tests, fixed before it could hit a real rip:** `(existingEpisodeNumbers | Measure-Object -Maximum).Maximum` returns a `Double` in Windows PowerShell 5.1 even when every input is an `Int32`. The episode tag was built with `"E{0:D2}" -f $nextEpisode`, and `.NET`'s `"D2"` format specifier only accepts integral types — a `Double` throws `FormatException` at runtime ("Format specifier was invalid"). This would have failed on exactly the disc-2-auto-detect case the whole feature exists for. Fixed with an explicit `[int](...)` cast around the `Measure-Object` result in both scripts.

**Files changed:**
- `rip-disc.ps1` — `$script:IsGenreSeries`/`$script:GenreLabel`/`$script:GenreFolder` computed after config load; genre-series branches added to `Get-TitleSummary`, the Ready-to-rip Type display, the final-output-dir routing, the `Write-Log "Type:"` line, `$contentType`, and a new Step 3 organize branch (episode numbering + move-up + Disc$Disc cleanup)
- `continue-rip.ps1` — identical changes, so resuming a genre-series rip at Step 3 organizes files the same way the original `rip-disc.ps1` run would have
- `README.md` — new "Documentary / genre series" usage section with examples, new directory-structure example, Feature Parity table row (PowerShell only, same as the other genre flags)

**Files deliberately NOT changed:**
- `RipDisc/*.cs` (C# implementation) — `-Documentary` and friends were never ported (see Feature Parity table); genre series inherits the same gap
- The `-Queue`/`-processQueue` entry hashtable in `rip-disc.ps1` (~line 1510) does not carry genre flags through at all — a **pre-existing** gap that affects every genre flag, not just this feature. Not fixed here (out of scope, and the boxset workflow described is sequential across sessions, not queued/concurrent) — worth a follow-up if genre flags + `-Queue` is ever actually used together.
- No extras-disc concept for genre series (matches plain `-Series`, which also has none) — if a documentary disc ships genuine bonus featurettes alongside episodes, they'll currently be numbered as episodes too. Left for the user's judgement (rip bonus content as a separate movie-mode disc, or move files manually afterward) rather than guessed at.

**Testing status:** `PSParser::Tokenize` reports 0 errors on both scripts; UTF-8 BOM confirmed intact on both. 12/12 logic unit tests pass against fixture files (not a real disc) covering: sequential numbering within a disc, auto-detected continuation across simulated sessions, variable episode-count-per-disc, explicit `-StartEpisode` override, filename-collision fallback, and the season-tag (`S01E0x`) variant — this is what caught the `Double`/`D2` bug above. **Not exercised against a real MakeMKV/HandBrake rip** — no disc was ripped this session.

**Episode names (added after the first round of review):**
Episodes are titled from the disc's own volume label, normalised through the existing `Clean-DiscName` (underscores → spaces, whitespace collapsed, title case) — so `WARMING_BY_THE_DEVILS_FIRE` becomes `Warming By The Devils Fire`. This is only safe when a disc holds exactly one episode; one label cannot name several files, so multi-episode discs fall back to numbering unless `-EpisodeNames` is supplied. `-EpisodeNames` always wins over the label. Filenames follow Jellyfin's documented `Series - S01E04 - Episode Name` pattern; unnamed episodes keep the original `<title>-E04.ext` shape.

Note the original premise of this feature — "every disc reports the same label, so discs can't be told apart" — **does not hold for the Blues box set**. Each disc self-identifies (`SOUL_OF_A_MAN`, `ROAD_TO_MEMPHIS`, `WARMING_BY_THE_DEVILS_FIRE`, `Godfathers and Sons`), which is exactly what makes label-based naming work. Worth revisiting whether auto-detect-by-label should also drive disc ordering.

**Two bugs found by the committed tests:**
1. `Resolve-EpisodeNames` returned a one-element array for a single-episode disc, which PowerShell unrolls to a bare string on return — the caller's `[0]` then indexed the *string* and produced `"W"` as the episode name. Every single-episode disc (i.e. the entire box set) would have been misnamed. Fixed with the unary comma (`return ,$names`) plus `@()` at the call site.
2. The episode-detection regex `-E(\d+)\.` demanded a dot immediately after the digits. Named episodes (`Title - E04 - Name.mp4`) never matched, so cross-session numbering silently restarted at 1 and overwrote earlier discs. Widened to `(?:^|[-\s])E(\d+)(?=\s|\.|$)`.

**Testing status:** `PSParser::Tokenize` reports 0 errors on both scripts. **25/25 logic tests pass** via `.\tests\Test-EpisodeNaming.ps1` — a committed, re-runnable suite that lifts the real function bodies and the real regex out of the scripts with the PowerShell AST parser, so the tests cannot drift from the shipped code without failing. It also asserts `Get-EpisodeFileName` and both detection patterns are identical across `rip-disc.ps1` and `continue-rip.ps1`. **Still not exercised against a real MakeMKV/HandBrake rip.**

Earlier entries in this file claim "8/8" and "12/12 logic unit tests pass" with no committed tests — those were scratch tests, run and discarded. `tests/` is the first re-runnable suite in this repo; prefer adding to it over ad-hoc verification.

**Work In Progress:**
- None. PR #116 was taken out of draft and **merged** as squash commit `537b6b7`, at the user's explicit instruction, **without real-disc validation** — see the disclosure comment on the PR. PR #117 (session-log reminder) merged first as `1ef0a44`; the two collided only on `CHANGELOG.md` ordering and both sets of entries were kept.
- The caveat that kept #116 in draft still stands, it is simply now on `main`: MakeMKV, HandBrake, real drive enumeration and `Get-DiscVolumeLabel` against actual hardware are all unexercised. The 25/25 tests cover extracted logic against fixtures, not an end-to-end rip.

**Outstanding Work for Future Sessions:**
- Real-world validation: rip an actual disc from the Martin Scorsese Presents the Blues boxset with `-Documentary -Series -Disc 1`, confirm the output layout matches the README example, then rip a second disc in a later session and confirm auto-numbering continues correctly
- Consider whether the `-Queue` genre-flag gap (noted above) is worth fixing generally
- Verify the CSS fix (PR #105) on a real rip — still parse-checked only, never runtime-tested
- Investigate why H: (`GP75N 1.01 K0MMB391933`) cannot authenticate — check region code and RPC setting
- Validate stuck sector detection on a genuinely damaged disc
- The C# port items have now been carried since 2026-02-16 (over six months) with no commit touching `RipDisc/*.cs` — still unresolved whether to formally abandon this
- Stray `nul` files at repo root and `.claude\nul` — still not cleaned up, still low priority
