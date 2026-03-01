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
