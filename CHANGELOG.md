# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/).

## 2026-08-29 - eBay Sold-Price Check

### Added
- New `-CheckEbayPrice` switch on `rip-disc.ps1` - prints a clickable eBay UK sold-listings search URL for the ripped title in the FILE SUMMARY, so you can check what the physical disc might be worth. Direct user request, ported straight over from the same feature added to the sibling `ripaudio` project earlier the same day.
- New `Get-EbaySoldListingsUrl` helper builds `https://www.ebay.co.uk/sch/i.html` with `_nkw=<title> <DVD|Blu-ray>[ Season N]` (URL-encoded via `[System.Web.HttpUtility]::UrlEncode`, newly loaded for this), `_sacat=0`, `_from=R40`, `LH_BIN=1` (Buy It Now only), `LH_ItemCondition=4` (Very Good or better), `LH_PrefLoc=1` (UK only), `rt=nc`, `LH_Sold=1` (sold listings only) - the same filter combination as `ripaudio`'s version. The format word (`DVD` vs `Blu-ray`) follows `-Bluray`; `-Series -Season N` appends `Season N` to the query.
- Off by default - a convenience for deciding what to do with a physical disc after ripping it, not part of the rip pipeline itself. Only wired into `rip-disc.ps1`'s main completion path, not the `-Queue` "job queued" summary (nothing has been ripped yet at that point) and not `continue-rip.ps1` (not carried over in this pass).
- README.md: new Features bullet, `-checkEbayPrice` usage-block entry and example, and a Feature Parity table row (PowerShell only - not ported to C#).

**Testing status:** `Parser::ParseFile` reports 0 errors; UTF-8 BOM confirmed intact by inspecting raw file bytes; all added lines confirmed ASCII-only. The URL-construction logic was verified standalone against three cases (a Blu-ray movie, a series with a season number, a plain DVD movie), producing correctly `+`-encoded query strings matching the `ripaudio` version's format. **Not exercised through an actual `rip-disc.ps1` run** - nobody has seen the FILE SUMMARY line render or opened the resulting URL in a browser to confirm eBay honours all four filters.

## 2026-08-26 (yet again, once more) - Sanitize Title Before Building Any Path

### Fixed
- Live incident: `rip-disc.ps1 -title "The Arena Hawaii 05/06 Highlight Reel - ASL" -Documentary` failed continuously with `Add-Content : Could not find a part of the path 'C:\Video\logs\The Arena Hawaii 05\06 Highlight Reel - ASL_disc1_...log'`. Root cause: `$title` was used raw everywhere a path or filename got built, and Windows treats `/` exactly like `\` — a `/` inside the title silently turned one intended folder into two nested ones (`...\The Arena Hawaii 05\06 Highlight Reel - ASL\`), and broke `Write-Log`'s `Add-Content` outright since only the outer directory had been created.
- Not fatal to the rip itself — MakeMKV and HandBrake both completed regardless, and `New-Item -ItemType Directory` creates the full nested path automatically — but the final output landed in the wrong nested location, file-prefixing in Step 3 would have used only the innermost folder name as the prefix (silently dropping everything before the `/`), and no session log was written for the run at all.
- A `$safeTitle` sanitizer (`-replace '[\\/:*?"<>|]', '_'`) already existed in both scripts, but only for the recovery-script filename — defined far too late to help `$finalOutputDir`, `$makemkvOutputDir`, or the log file path, all of which are built earlier. Moved the sanitization to the top of the `CONFIGURATION` section in both scripts, immediately after `$title` is finalized, and switched every path/filename-building use of `$title` to `$safeTitle`: `$makemkvOutputDir` (all 3 branches), every `$finalOutputDir` branch (genre, genre-series, series, Blu-ray, DVD), `$script:LogFile`, the series-mode `$prefix`, and the extras-mode prefix-matching block. Display-only uses of `$title` (console output, log message text, window title, TMDb lookups) are deliberately untouched, so the user still sees their real title everywhere it's shown, not the sanitized form.
- Movie-mode's main-feature-disc prefixing needed no change — it already reads the prefix back from the real directory (`(Get-Item $finalOutputDir).Name`) rather than reconstructing it from `$title`, so it self-corrects now that `$finalOutputDir` itself is built from `$safeTitle`.

**Testing status:** `Parser::ParseFile` reports 0 errors on both scripts; UTF-8 BOM confirmed intact by inspecting raw file bytes. Full existing test suite re-run: 95/95 passing, no regressions. Not yet re-exercised against a real disc with a slash-containing title — the live incident's rip was left running rather than interrupted, so this fix will first prove itself on the next affected title.

## 2026-08-26 (continued yet again) - Widen the File-Lock Retry Window on Rename/Move

### Fixed
- Live incident: `rip-disc.ps1`'s Step 3 organize failed on "Coast Video Magazine No. 2" — 3 of 4 encoded files renamed fine, but the 4th (`D3_t03.mp4`, the most recently written) stayed locked through the full retry window and the step aborted with `The process cannot access the file because it is being used by another process.` MakeMKV and HandBrake had both already completed successfully; nothing was lost, the file just never got renamed.
- Root cause: the existing safety margin (a 3s post-encode pause, then 5 rename attempts 3s apart) totals 15s, which wasn't enough here — the pattern (only the last-written file affected) points at something like Windows Defender's real-time scan grabbing the file the moment HandBrake releases it, which can hold it well past 15s under load.
- Widened `$maxRetries`/`$retryDelay` from `5`/`3` to `10`/`5` (50s worst case instead of 15s) at all 12 call sites across both scripts — 6 rename/move sites each in `rip-disc.ps1` and `continue-rip.ps1`, all copy-pasted from the same original pattern (PR #24) and updated uniformly rather than special-casing the one that failed tonight.
- Not a structural fix — the 12-way duplication itself is untouched (extracting a shared retry helper would be a larger, separate change); this is a targeted widening of the existing pattern.

**Testing status:** `Parser::ParseFile` reports 0 errors on both scripts; UTF-8 BOM confirmed intact (verified by inspecting raw file bytes, not just re-decoding — an early draft of this fix accidentally wrote a second BOM into both files by decoding without stripping the first one, caught before commit). Full existing test suite re-run: 95/95 passing across all 5 files, no regressions. Not yet re-exercised against a real disc under AV load — the underlying trigger (Defender scanning a freshly-closed large file) can't be reliably reproduced on demand.

## 2026-08-26 (continued again) - Test Coverage and C# Port for the Drive-Query Hang Fixes

### Added
- `tests/Test-DriveQueryTimeout.ps1` — 16 tests covering the two genuinely pure/reusable pieces of the drive-query hang fixes below, extracted into named functions specifically so they could be tested (they were previously inline in the main script body):
  - `Select-MatchedDrive` (new function) — the duplicate-drive-index selection logic from the earlier fix, tested against a fixture reproducing the exact live scenario (two `DRV:` entries sharing the same index, one by drive letter, one by raw device path)
  - `Wait-ProcessWithTimeout` (new function) — the process-timeout-and-kill mechanic, tested against real short-lived child processes (a fast-exiting one, and one simulating a hang via `Start-Sleep`) rather than mocked, since it's inherently a real-process integration point
  - `rip-disc.ps1`'s inline drive-matching and timeout-polling code now calls these two functions instead of duplicating the logic inline - a behaviour-preserving refactor, verified by re-running the full existing 79-test suite (all still passing) before and after
- Ported the equivalent of the "misleading no-disc-detected message" fix (see the 2026-08-26 entry below, PR #123) to the C# implementation (`RipDisc/RipDisc/RipDiscApplication.cs`) - the exact same `.Contains("0 titles")` overmatch bug existed there too, in `AnalyzeMakeMKVNoFilesError`. Added the same higher-priority device-disconnect check (`STATUS_DEVICE_NOT_CONNECTED` / "does not exist") ahead of the narrowed "no valid title" check, in both `AnalyzeMakeMKVError` and `AnalyzeMakeMKVNoFilesError`
- Added a 4-hour safety-net timeout to the C# implementation's MakeMKV rip call, via a new optional `timeoutSeconds`/`timeoutStepLabel` parameter pair on `ExecuteProcess` (defaults to unbounded - the two HandBrake call sites are unaffected). This is *not* a literal port of the PowerShell drive-query timeout: the C# app has no equivalent short drive-enumeration call to begin with (it goes straight to `dev:{driveLetter}` or `disc:{DriveIndex}`, unlike the PowerShell version's `-r info disc:9999` step), so there was nothing to attach a short timeout to. Instead this guards the one call site that actually shares the underlying risk - a hung/disconnected drive during the real rip - with a timeout generous enough to never interrupt a legitimate rip
- README Feature Parity table updated to reflect both ports accurately, including a new row distinguishing the two different hang-protection *mechanisms* (PowerShell's existing stuck-sector pattern detection vs. C#'s new flat safety-net timeout) rather than claiming a false 1:1 match

## 2026-08-26 (continued) - Drive-Query Hang: Root Cause, Timeout, and Diagnostics

### Fixed
- `Get-DiscVolumeLabel`'s `Get-CimInstance Win32_CDROMDrive` call in `rip-disc.ps1` could hang for minutes (found live: a USB DVD drive that had started disconnecting mid-rip made WMI's drive enumeration stall) even when the run targeted a completely different, healthy drive - this call enumerates every CD-ROM drive on the system, not just the target one. Added `-OperationTimeoutSec 5`, bounding it to a few seconds; the existing `catch` already treats a failed lookup as non-fatal (episode-naming convenience only), so a timeout now degrades gracefully instead of hanging the whole script
- The "MakeMKV drives:" listing and the "Using disc:N (...)" confirmation could show a drive's name doubled up (e.g. "GP75N ... GP75N ...") and mark two different listed drives with `<--` at once. Root cause: a drive that reconnects mid-session can appear in MakeMKV's own drive list twice under the same `disc:N` index - once keyed by drive letter, once by raw device path (e.g. `\Device\CdRom3`) - so matching by index alone silently produced a 2-element array, and PowerShell's string interpolation joined both `.Name` values with a space. Both the confirmation-line lookup and the `<--` marker now also require the drive letter to match, not just the index
- The real cause of the live hang the two fixes above were chasing: `& $makemkvconPath -r info disc:9999 2>&1` had no timeout mechanism at all. MakeMKV's own drive-enumeration query can hang indefinitely when a physical drive is malfunctioning (confirmed live: 5+ minutes and counting on a USB DVD drive that had started dropping connection) - this happens *before* either of the fixes above ever run, so neither could help. Replaced the bare `&` call with a real `System.Diagnostics.Process`, polled with a hard timeout; on timeout the process is killed and a clear error is shown instead of hanging forever. Verified both the timeout-and-kill path and the normal-output path against synthetic child processes before shipping
  - Follow-up, same day: the initial 30s timeout was too short - live evidence showed a legitimate (not stuck) query on a USB drive spinning up from idle taking ~30-35s to succeed on its own, so a 30s cutoff was killing queries that would otherwise have worked. Raised to 60s
  - The timeout error message was also too thin (no log file exists yet at this point in the script - logging only starts once a drive is identified - and no retry guidance was given). It now explicitly says why there's no log for this failure, checks for and reports any *other* `makemkvcon`/`makemkvcon64` process still running (a Ctrl+C'd earlier rip does not kill its own MakeMKV child process, which can then hold the drive exclusively and cause this exact symptom on the next attempt), and lists concrete retry options (`Stop-Process` the leftover, just retry, or use `-DriveIndex` to skip the lookup entirely)

**Testing status:** a subsequent rip attempt against the same drive completed successfully after the 60s/diagnostics change (#126) shipped - the hang did not recur. That is one successful run after several same-session fix iterations on live, flaky hardware; treat as encouraging, not as confirmation the hang is fully resolved. See `CLAUDE.md` session notes for the full incident chain and what still needs follow-up.

## 2026-08-26

### Added
- On failure, `rip-disc.ps1` now prints a ready-to-paste `continue-rip.ps1` command under `--- RETRY WITH continue-rip.ps1 ---`, reformatted from this run's own inputs
  - New `Get-ContinueRipCommand` function maps the earliest incomplete step to `continue-rip.ps1`'s `-FromStep` (`handbrake`/`organize`/`open`) and carries over `-title`, `-Series`/`-Season`/`-Disc`, genre flags, `-StartEpisode`, `-EpisodeNames`, `-NoSound`, and (only when explicitly passed) `-OutputDrive`; values matching `continue-rip.ps1`'s own defaults are omitted to keep the command short
  - Returns `$null`, and prints nothing, when Step 1 (the MakeMKV rip) itself never completed — `continue-rip.ps1` resumes AFTER that step, so there is nothing for it to work with
  - Titles/episode names containing a literal double quote are escaped (`` ` " ``) so the printed command still parses as a single argument when pasted back into PowerShell

### Fixed
- Genre-series episode naming (and the "MakeMKV drives:" listing) could use a stale disc label after swapping discs in the same drive within the drive-lookup cache's 5-minute TTL — the cached MakeMKV `DRV:` line, not the disc actually in the drive, was shown and used to name episodes
  - The target drive's entry now prefers a live, per-drive-letter Windows volume-label query (`Get-DiscVolumeLabel`, already used elsewhere as a fallback) over the cached MakeMKV name, so a disc swap is reflected immediately without re-enumerating every drive
  - Scoped to the one matched drive only, so the existing cache (and the slow full `disc:9999` re-query it exists to avoid) is untouched for every other drive
  - Only applies when `-DriveIndex` is not used — that path has no drive letter to query and already had no label available (documented pre-existing limitation)

### Changed
- Dropped a redundant `-and $DriveIndex -lt 0` condition from the disc-label live-query guard added above — the surrounding branch only ever runs when `$DriveIndex` is unset, so the check was always true and added nothing (cosmetic follow-up, same run's PR review)

### Added
- `tests/Test-ContinueRipCommand.ps1` — 11 tests for `Get-ContinueRipCommand`, promoting the ad hoc verification from the retry-suggestion work above into a committed, re-runnable suite (AST-extracted from `rip-disc.ps1`, same technique as the other `tests/*.ps1` files). Covers step-to-`-FromStep` mapping, the step-1-unresolved `$null` case, default-value omission, non-default value carry-through and ordering, and quote escaping in `-title`/`-EpisodeNames`

## 2026-08-25

### Added
- `-NoSound` parameter in `rip-disc.ps1` and `continue-rip.ps1` — skips the completion fanfare (`[Console]::Beep` melody) played at the end of a run (both the normal completion path and, in `rip-disc.ps1`, the `-Queue` completion path)
- `-NoEject` parameter in `rip-disc.ps1` — skips ejecting the disc after the MakeMKV rip completes, leaving it in the drive
  - `continue-rip.ps1` also accepts `-NoEject` for command-line compatibility (so a failed `rip-disc.ps1` command can be pasted there unchanged), but ignores it — it resumes after the rip step and never ejects, the same treatment already given to `-Drive`/`-DriveIndex`

**Testing status:** Parse-checked only (`[System.Management.Automation.Language.Parser]::ParseFile` reports 0 errors on both scripts) and the UTF-8 BOM was verified intact after editing. Not runtime-tested against a real disc.

## 2026-08-24

### Added
- Closing session-log reminder in `rip-disc.ps1` and `continue-rip.ps1` — both scripts now end with a `--- SESSION LOG ---` section pointing at the log they just wrote
  - Two new functions per script, `Format-TerminalLink` and `Show-LogFileReminder`, called as the last output on the success path and from `Stop-WithError` on the failure path
  - The literal path is always printed so it can be copied or pasted regardless of terminal; the containing folder and a `notepad "<path>"` hint are printed alongside it
  - `Format-TerminalLink` wraps the path in an OSC 8 hyperlink so it is clickable, but only when the host is known to support one (`$env:WT_SESSION` for Windows Terminal, `$env:TERM_PROGRAM -eq 'vscode'`). Detection is opt-in rather than assumed because legacy conhost renders the escape sequence as visible garbage; unrecognised hosts get plain text
  - ESC is built as `[char]27` — PS 5.1 has no `` "`e" `` escape
  - The log path is chosen up front, before anything is written to it, so when the file does not exist the reminder says so plainly rather than pointing at a missing file
  - Both functions are duplicated verbatim across the two scripts, matching the existing pattern in this repo
- `tests/Test-LogFileReminder.ps1` — 16 tests covering terminal capability detection, `file://` URI construction (including percent-encoded spaces), and the reminder's output
  - Extracts the real function bodies from both scripts via the PowerShell AST parser, so it exercises shipped code rather than a copy, and asserts the two scripts' copies stay identical
- Documentary / genre series mode in `rip-disc.ps1` and `continue-rip.ps1` — for multi-disc box sets that belong in a genre folder rather than `Series\`
  - Combine `-Series` with any genre flag (`-Documentary`, `-Tutorial`, `-Fitness`, `-Music`, `-Surf`) to activate it; plain `-Series` and plain genre flags are unaffected
  - Reuses `-Series`' existing per-disc isolation and composite mega-file detection rather than adding a parallel system
  - Every MKV on the disc is treated as an episode of equal standing (no Feature/extras split) and renamed `<title>-E##.mp4` (or `<title>-S##E##.mp4` with `-Season`)
  - Episodes are moved out of the per-disc `Disc$Disc` subfolder into a single shared title folder, and the emptied `Disc$Disc` folder is removed — unlike plain `-Series` mode, which leaves episodes nested in `Disc$Disc` permanently
  - Episode numbering carries across sessions automatically: `-StartEpisode` is now optional and, when omitted, is auto-detected from the highest existing episode file already in the destination folder
  - `-StartEpisode` still works as an explicit override
- Episode names for genre series, so a box set lands as `<title> - S01E04 - <Episode Name>.mp4` (Jellyfin's documented pattern) rather than bare numbering
  - Titles are taken from the disc's own volume label, normalised through the existing `Clean-DiscName` (underscores to spaces, whitespace collapsed, title case) — `WARMING_BY_THE_DEVILS_FIRE` becomes `Warming By The Devils Fire`
  - Only applied when a disc holds exactly one episode; one label cannot name several files. Generic labels (`DVD_VIDEO`, `UNTITLED`, and similar) are rejected
  - New `-EpisodeNames` parameter in both scripts always overrides the label, and is the only option for a disc yielding multiple episodes. Episodes with no name keep the existing `<title>-E##.ext` shape
  - `Get-DiscVolumeLabel` falls back to Windows for the drive letter when MakeMKV reports no label (reproducible on the USB DVD units here). No label is available under `-DriveIndex` (the drive list is skipped) or in `continue-rip.ps1` (it never reads the disc), so `-EpisodeNames` is required there
- `tests/Test-EpisodeNaming.ps1` — the first committed, re-runnable test suite in this repo (25 tests). It lifts the real function bodies and the real regex out of both scripts with the PowerShell AST parser, so the tests cannot drift from the shipped code without failing, and asserts the shared logic is identical across `rip-disc.ps1` and `continue-rip.ps1`

### Fixed
- `(... | Measure-Object -Maximum).Maximum` returns a `Double` in Windows PowerShell 5.1 even for all-integer input, which throws `FormatException` against the `"D2"` format specifier used to build episode numbers — caught by the new logic unit tests before it could hit a real rip; fixed with an explicit `[int]` cast in both scripts
- `Resolve-EpisodeNames` returned a one-element array for a single-episode disc, which PowerShell unrolls to a bare string on return — the caller's `[0]` then indexed the *string* and yielded its first character as the episode name, misnaming every single-episode disc. Fixed with the unary comma (`return ,$names`) and `@()` at the call site
- The episode-detection regex `-E(\d+)\.` required a dot immediately after the digits, so named episodes (`Title - E04 - Name.mp4`) never matched and cross-session numbering silently restarted at 1, overwriting earlier discs. Widened to `(?:^|[-\s])E(\d+)(?=\s|\.|$)` in both scripts

### Changed
- The mid-run `Log file:` line in `Stop-WithError` moved from above the error banner to after it, in both scripts — on a failure the log is what the user needs, and printing it first meant it scrolled away behind the recovery guidance
- README "Session logging" feature bullet extended to mention the clickable log path shown at the end of every run

## 2026-08-24 (continued) - Null output path and wrong resume-step hint

### Fixed
- `$finalOutputDir` could end up `$null`/empty in both `rip-disc.ps1` and `continue-rip.ps1`, producing raw `Test-Path`/`Join-Path` parameter-binding exceptions and a blank `Output folder :` line instead of a clear error
  - A malformed or empty `-OutputDrive` can make one of the `Join-Path` calls that build `$finalOutputDir` emit a *non-terminating* error and return nothing while the script continues regardless — e.g. a provider-qualified-looking path such as `F::\...` fails with `Cannot find a provider with the name 'F'` rather than throwing
  - Fixed with a construction-time validation immediately after `$finalOutputDir` is built in both scripts: it must match `^[A-Za-z]:\\`, or the script stops immediately with a clear message showing what it resolved to
  - Defensive `-not [string]::IsNullOrWhiteSpace(...)` guards also added at the two original crash sites in `continue-rip.ps1` (the prerequisite check and the already-encoded lookup) as a second line of defense
- Both scripts now fail fast when the destination drive itself is missing, instead of running the prerequisite checks (and, previously, starting an encode) against an unusable path
  - Reuses the existing `Test-DriveReady` helper, called right after `$finalOutputDir` is validated — before `rip-disc.ps1` even starts ripping the disc in Step 1
  - `Test-DriveReady` itself now gives a clear "output path is empty" message when handed a null/empty path, instead of the misleading `Could not determine drive letter from path:` with nothing after the colon
- `continue-rip.ps1`'s "To pick up from here" resume hint suggested `-FromStep 4` after a **Step 2** failure, which would have skipped the encode and organize steps entirely
  - Root cause: `Sort-Object Number` silently sorts **descending** on Windows PowerShell 5.1 when the pipeline objects are `[hashtable]` (as `$script:AllSteps` entries are) rather than `[PSCustomObject]` — the bare property-name binder doesn't resolve through the same adapter that dot-notation member access uses, so it fell back to a different (and here, reversed) ordering
  - Fixed by sorting on a script block (`Sort-Object { $_.Number }`) instead, which reads the value directly and sorts correctly regardless of object type
  - `rip-disc.ps1` does not have this defect — it has no resume-hint feature to begin with (`-FromStep` only exists in `continue-rip.ps1`)

### Added
- `tests/Test-ContinuePathAndResume.ps1` — 27 tests covering both fixes above: the `$finalOutputDir` validation condition (extracted from the real source text) against the exact failure shapes seen in this incident plus several adjacent malformed-path cases, `Test-DriveReady`'s new empty-input handling in both scripts, and the resume-hint regression (asserted against the real `Show-StepsSummary`/`Get-RemainingSteps`/`Get-Step` function bodies)

## 2026-08-17

### Added
- Blu-ray mode guard in `rip-disc.ps1` and `continue-rip.ps1` — catches a missing `-Bluray` flag before encoding starts
  - Ripping a Blu-ray without `-Bluray` takes the DVD subtitle branch (`--all-subtitles --subtitle-burned=none`), and Blu-ray PGS tracks get burned into the picture anyway despite that flag (the reason the `-Bluray` branch exists — see #67)
  - Burned-in subtitles cannot be removed afterwards, and `rip-disc.ps1` deletes the source MKVs as soon as Step 2 finishes, so the mistake was unrecoverable without re-ripping the disc
  - Detection: any single ripped MKV larger than 8.5 GB. A title cannot be bigger than the disc it came from, and DVD-9 tops out at 8.5 GB, so anything above that is proof of a Blu-ray source
  - Deliberately checks each file rather than their total — MakeMKV often emits the same feature as more than one title, so a DVD's files can legitimately sum well past 8.5 GB (a 3×4 GB DVD rip must not trigger this)
  - Costs nothing: the files are already on disk and their sizes are already read for the existing log lines
  - On a hit, offers: enable Blu-ray subtitle handling (default), continue in DVD mode anyway, or abort. Abort leaves the MKVs in place and prints the `continue-rip.ps1 -FromStep 2 -Bluray` command to resume without re-ripping
  - The default — including an unanswered or piped-EOF prompt — is the corrective, non-destructive option, so it cannot repeat the pattern where a blank line was read as consent to start a destructive encode
  - In `rip-disc.ps1` the check sits before the queue block, so a queued job carries the corrected flag through to encoding
  - `continue-rip.ps1` honours `-Yes` by enabling Blu-ray handling automatically rather than prompting
  - Output directory routing is deliberately left alone: `$finalOutputDir` is resolved before the rip and the folder already exists, so only subtitle handling changes. This is logged and shown

### Fixed
- Open-directory step no longer fails on titles containing spaces, in both `rip-disc.ps1` and `continue-rip.ps1`
  - Both scripts called `Start-Process explorer.exe -ArgumentList $directoryToOpen` with the path unquoted
  - `Start-Process` joins `-ArgumentList` on spaces without quoting the elements, so `C:\Video\Who Framed Roger Rabbit\Disc1` reached explorer.exe as three separate arguments and only the first token was treated as a path
  - Titles with spaces are the norm here, so this fired on most rips — impact was limited to the final "open the folder" convenience step, not the rip itself
  - The path is now wrapped in embedded quotes, with `TrimEnd('\')` so a trailing backslash cannot escape the closing quote
  - Same defect class as the ripaudio fix of the same date, where it was more severe: there it broke the automatic handoff to `search-metadata.ps1`

## 2026-08-11

### Fixed
- `continue-rip.ps1` Step 3 no longer renames files in the wrong directory when the output folder cannot be entered (#107)
  - Step 3 ran a bare, unchecked `cd $finalOutputDir`, and the entire organize block renames and moves files relative to the *current* directory
  - If that `cd` failed, the script carried on and renamed whatever happened to be in the directory it was launched from — this actually hit the repo working directory
  - Now `Set-Location -LiteralPath $finalOutputDir -ErrorAction Stop` inside a try/catch, plus a post-check comparing `(Get-Location).Path` against the target before any renaming takes place
- `rip-disc.ps1` Step 3 gets the same working-directory guard as `continue-rip.ps1` above (#109)
  - Same bare, unchecked `cd $finalOutputDir` and the same class of bug: the organize block renames, moves and deletes files relative to the *current* directory
  - `Set-Location -LiteralPath $finalOutputDir -ErrorAction Stop` inside a try/catch routing to `Stop-WithError`, plus a guard comparing `(Get-Location).Path` against the target before any renaming
  - Preventative — this fix has not fired against a real failure; the actual incident (every file in the repo renamed with a leading `-`) went through `continue-rip.ps1`'s Step 3, fixed in #107 above

### Changed
- `continue-rip.ps1` is now usable without knowing the arguments up front (#107)
  - `title` and `FromStep` are optional; omitting them shows a step menu rather than failing on a mandatory-parameter prompt
  - `-Drive` and `-DriveIndex` are accepted and ignored, so a failed `rip-disc.ps1` command line can be pasted straight into `continue-rip.ps1` without editing
  - Every step is always listed with its prerequisites, and the chosen step can be switched at the prompt
  - Prerequisite checks extracted into `Test-StepPrerequisites`; `Stop-Prerequisite` became `Write-PrerequisiteFailure`, which returns `$false` instead of calling `exit 1`, so a failed check re-offers the step list instead of killing the session
  - `-Yes` still exits 1 on a failed prerequisite, preserving non-interactive behaviour
- `continue-rip.ps1` Step 2 skips MKVs that already have a matching MP4 instead of re-encoding them (#110)
  - Resuming a rip usually means encoding stopped part way through, so redoing finished files wasted hours
  - New `-Force` switch (alias `-ReEncode`) restores the old encode-everything behaviour
  - Each skipped file is printed and logged with the existing MP4's size; skipping assumes those MP4s are complete — it checks existence, not integrity
  - Reports "Nothing to encode" and continues to Step 3 when every MKV already has an MP4
  - Recovery script is now only generated when there is something to encode, built from the filtered list; `$recoveryScriptPath` starts `$null` with its deletion guarded so `Test-Path` is never handed a null path
  - The encode loop and its "file N of M" counters (console and log) are driven from the filtered list
  - The pre-flight prerequisite summary now reports which files will be skipped, or overwritten under `-Force`, and a count of files to encode this run

### Verified
- The 2026-08-10 eject fix (#106) confirmed working in production on its first real rips, under three concurrent sessions
  - Drive G: MakeMKV complete 06:18:54, ejected 06:19:00; drive H: complete 06:24:23, ejected 06:24:29 — 6 seconds each
  - No retry attempts and no timeouts, against 8 failures out of 10 rips under the same concurrency the previous afternoon

## 2026-08-10

### Fixed
- Disc eject no longer falsely reported as timed out while the disc is sitting in an open tray (#106)
  - The eject ran via `Start-Job` + `Wait-Job -Timeout 15`, which measured PowerShell child-process startup, not the eject
  - Under the CPU saturation this script creates itself (12 concurrent `HandBrakeCLI` encodes pinning the CPU at 100%), a `Start-Job` with a trivial `{ 1 }` body was measured at 18.0s, 25.7s and 33.2s — every one of them over the 15s deadline before the eject could report back
  - Explains why the failure rate climbed through the day and hit all drives equally: 2026-06 logs show zero eject timeouts, and on 2026-08-10 every rip up to 15:14 succeeded while 8 of the 10 after it failed, as concurrent encodes accumulated
  - The eject is now issued in-process as a direct `IOCTL_STORAGE_EJECT_MEDIA` device call — no child process, and it touches only the target drive, unlike the old `Shell.Application` verb which made Explorer re-enumerate every optical drive and could stall on a sibling drive mid-rip in a concurrent session
  - Success is confirmed by polling `System.IO.DriveInfo.IsReady` until the drive goes not-ready, rather than by trusting a call to return inside a deadline — it now reports what actually happened rather than how long a process took to spawn
  - Measured on drive H: at 100% CPU load: IOCTL 1309ms, eject confirmed after 267ms with zero poll iterations
  - The failure dialog is now reworded, since a genuine failure is no longer a timeout
- Stuck sector detection no longer aborts rips because of read errors from a different drive (#105)
  - `makemkvcon` enumerates every optical drive at startup, so CSS errors from an unrelated drive arrived before the rip began, tripped the watchdog after 5 lines and killed the rip roughly 2 seconds in
  - Offset errors now only count toward the stuck counter once the rip has actually started (`Saving N titles`, `Current progress`, `Current operation`, or `Title #`)
  - Errors are attributed to a drive by parsing `occurred while reading '<drive>' at offset '<n>'` and comparing against the target drive; errors naming another drive are still printed but never counted
  - Added a pre-rip escape hatch: 50 consecutive read errors before the rip starts also kills MakeMKV, so a disc that can never be authenticated cannot hang indefinitely
  - The `-DriveIndex` path has no drive list to compare against, so it still counts every error as before
- Stuck-sector kill no longer masked as success (#105)
  - `$makemkvExitCode = if ($wasKilledForStuck) { 0 }` forced exit code 0, skipping the entire error-analysis block so the user only ever saw the generic "No MKV files were created"
  - The exit code is now judged by what was salvaged: 0 only when a stuck kill left at least one MKV file, otherwise non-zero so error analysis runs
- CSS authentication failure detected without requiring "Failed to open disc" (#105)
  - The CSS branch was nested inside a `Failed to open disc` check, which the real-world failure never printed — only SCSI errors — so it never fired
  - `SCRAMBLED SECTOR WITHOUT AUTHENTICATION` is now matched anywhere in the captured output, names the offending drive when parseable, and notes it may be a different drive from the one being ripped
- Drive list cache no longer stores a failed enumeration (#105)
  - `makemkv-drive-cache.txt` stored raw `info disc:9999` output including `MSG:2003` error lines; a drive that errors may be missing or wrong in the `DRV:` list, and the bad mapping was then reused for the full 5-minute TTL
  - The cache is only written when the enumeration is error-free; otherwise the drive list is re-queried on the next run

### Added
- Target drive name (`$script:TargetDriveName`) recorded from the MakeMKV drive list so read errors can be attributed to a specific drive (#105)

## 2026-03-23

### Fixed
- Drive index mapping: replaced WMI `Win32_CDROMDrive` enumeration with direct MakeMKV `disc:N` iteration (#96)
  - WMI ordering did not match MakeMKV's internal disc index, causing rips from the wrong physical drive
  - Now iterates `disc:0`, `disc:1`, etc. via MakeMKV and matches against the requested drive letter
- Concurrent rip safety: detects running `makemkvcon` processes and skips their active `disc:N` indices (#98)
  - Prevents the drive lookup from interfering with any concurrent rip sessions
- Stuck sector detection: kills MakeMKV when 5 consecutive errors occur at the same byte offset (#99)
  - Prevents indefinite stalls on physically damaged discs
- Stuck sector detection PS 5.1 compatibility: rewrote `Thread`/`ConcurrentQueue` approach to direct `Process` execution (#100)
  - PS 5.1 does not support the threading primitives used in the original implementation
- Busy-drive guard removed: `-not $isBusy` check incorrectly blocked selecting drives that were in use by a previous rip attempt (#101)
  - A drive can be "busy" in the process list while still available for a new rip session
- Exit code crash: replaced cmd.exe wrapper (which returned exit code 2 due to path quoting) with direct `Process` object execution (#101)
  - cmd.exe quoting caused makemkvcon to fail with exit code 2; direct invocation passes arguments correctly

### Added
- Startup drive list: all MakeMKV-detected optical drives are displayed at launch with an arrow on the selected drive (#97)
  - Allows visual confirmation that the correct physical drive is selected before ripping begins

## 2026-03-04

### Fixed
- MakeMKV drive path: replaced broken `dev:\\.\<PNP-DeviceID>` with `disc:N` format using WMI enumeration index (#75)
  - USB drives returned USBSTOR device paths which MakeMKV doesn't understand
  - Now maps drive letter to MakeMKV disc index via `Win32_CDROMDrive` enumeration
  - Added drive wake-up (`Test-Path`) before WMI query and Step 1 to spin up dormant USB drives
  - Error output now lists all available drives when target not found

## 2026-03-01

### Added
- `-Bluray` disc type now routes output to `<OutputDrive>:\Bluray\<title>\` instead of `DVDs\<title>\` (#69)
- Queue entry includes `Bluray` flag for downstream processing (#69)
- MakeMKV rip output now streams to the console in real time (#72)

### Fixed
- File rename no longer double-prefixes when MakeMKV filename already contains the title with underscore separator (#70, #71)
  - `Southpaw_t01.mp4` was becoming `Southpaw-Southpaw_t01.mp4`; now correctly becomes `Southpaw-t01.mp4`
  - Applied to all prefix paths: Disc 1 main feature, Extras disc, and Disc 2+ special features

## 2026-02-23

### Added
- Auto-discovery of disc metadata when `-title` is omitted (#57)
  - Reads disc info via MakeMKV info mode (disc name, type, title count)
  - Cleans disc name (strips suffixes, extracts season/disc hints, title-cases)
  - Searches TMDb API for canonical title and media type
  - Auto-populates `-title`, `-Bluray`, `-Series`, `-Season`, `-Disc`
  - Interactive confirmation prompt (Accept / Edit / Abort)
  - Falls back to manual input if disc name is too generic or TMDb unavailable
- Blu-ray format auto-detection from disc type even when `-title` is provided (#57)
- Disc format shown in "Ready to Rip" confirmation display (#57)

### Added (cont.)
- "Buy me a coffee" link shown after successful completion in both scripts (#60)

### Changed
- `-title` parameter is now optional (was mandatory) — defaults to auto-discovery (#57)
- `$makemkvconPath` moved earlier in script to support discovery functions (#57)
- Skip disc query entirely when `-title` is provided — `makemkvcon info` is too slow (30-60s) for just Blu-ray detection (#64)
- Discovery mode uses `disc:0` (MakeMKV auto-find) via separate `$discoverySource` variable; rip step respects `-Drive` parameter (#60, #65)

### Fixed
- `Get-DiscInfo` regex parsing: removed `$` anchors that fail on Windows `\r` line endings, added `.Trim()`, skip `ErrorRecord` objects from stderr (#60)
- `-Drive` parameter was being ignored — `disc:0` was always used instead of `dev:$driveLetter` (#65)
- `$discType` variable collision with `$script:DiscType` caused "Disc Format: Main Feature" instead of actual disc type — renamed to `$discTypeLabel` (#65)
- Blu-ray PGS subtitles no longer burned in — uses `--subtitle scan --subtitle-burned` to only burn forced/foreign-language subs (#67)

## 2026-02-22

### Added
- `-Music` disc type switch — outputs to `<OutputDrive>:\Music\<title>\` (#45)

### Fixed
- Disc 1 non-feature move to extras now uses `Get-UniqueFilePath` to avoid silent filename collisions during concurrent Disc 1 + Extras rips (#45)
- Series mode: `Remove-Item` on empty Disc subdirectory no longer fails when PowerShell's working directory is inside it (#53)
- Composite mega-file detection now uses sum-based heuristic (70-130% of sum of other files) instead of 2x second-largest, which failed when one episode was much longer than others (#54)

## 2026-02-18

### Added
- `-Tutorial`, `-Fitness`, and `-Surf` disc type switches — output to `E:\Tutorials\`, `E:\Fitness\`, `E:\Surf\` (#40, #44)
- Extras disc: encode directly into `extras\` subdirectory instead of encoding then moving (#43)
- Extras disc: files prefixed with title only, no `-extras` or `-Special Features` in name (#42)
- Empty parent directory cleanup after temp directory removal (#39)

### Fixed
- Series concurrent disc rename conflicts — use per-disc subdirs for encoding isolation (#41)
- Add UTF-8 BOM to scripts for PowerShell 5.1 compatibility — fixes parse errors on nested string expressions

## 2026-02-16

### Added
- `-StartEpisode` parameter for multi-disc season episode numbering offset (#32)
- Jellyfin episode naming format for series (`Title-S01E01.mp4`) (#31)
- Composite mega-file detection — skips all-in-one files during series encoding (#31)
- Disc eject retry — retries once after 2-second delay before giving up (#36)
- Windows dialog popup on eject timeout showing title and drive letter (#37)
- Triumphant completion fanfare in `rip-disc.ps1` for normal and queue modes (#37)
- Completion fanfare in `continue-rip.ps1` for parity (#38)

## 2026-02-09

### Fixed
- Disc 1 temp directory no longer collides with Disc 2+ during concurrent rips (#24)
- File lock retry messages now display correct filename in `ForEach-Object` blocks (#24)

## 2026-02-01

### Added
- `continue-rip.ps1` script for resuming failed rips from any step (#21)
- Blu-ray subtitle fallback — tries with subtitles first, retries without on PGS failure (#23)

### Changed
- Blu-ray mode now attempts subtitles before falling back, instead of skipping entirely (#23)

## 2026-01-29

### Added
- `-Bluray` flag for Blu-ray disc handling with PGS subtitle skip (#18)

## 2026-01-28

### Added
- `-Queue` flag to defer HandBrake encoding to a shared queue file (#17)
- `-ProcessQueue` flag to process all queued encoding jobs sequentially (#17)
- File locking for concurrent queue writes (#17)

## 2026-01-19

### Added
- `-OutputDrive` parameter to make output drive configurable (#6)
- Concurrent disc ripping with isolated temp directories per disc (#7)
- Window title management showing film name and status suffixes (#7)
- Extras disc window title uses lowercase `-extras` format (#8)

## Earlier Changes

### Added
- `-Documentary` flag for documentary output path (`E:\Documentaries\`) (#20)
- `-Extras` switch for extras-only disc ripping (#15)
- Series title validation warnings for misplaced metadata (#14)
- C# console application port (`RipDisc/`) (#14)
- Safety check for suspiciously small encoded files (#12)
- Console close button protection during rip (#11)
- Interactive prompt for existing MakeMKV output files (#10)
- HandBrake recovery script generation before encoding (#27)
- Drive readiness checks before operations (#6)
- Session logging to `C:\Video\logs\` (#9)
- Step tracking with completion summary (#4)
- `-Season` parameter for TV series numbering (#1)
- Automatic disc ejection after rip

### Fixed
- Directory creation failures now fatal instead of silently continuing (#29)
- Disc eject timeout prevents script from hanging (#28)
