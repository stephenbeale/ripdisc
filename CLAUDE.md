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

**Work In Progress:**
- None — all PRs merged, working tree clean

**Outstanding Work for Future Sessions:**
- Port missing features to C# implementation (see Feature Parity table in README)
