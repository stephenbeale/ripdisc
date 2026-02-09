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
