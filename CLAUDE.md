# RipDisc Project

PowerShell scripts for automated DVD and Blu-ray disc ripping using MakeMKV and HandBrake.

## Git Workflow

When the user says **"make a workflow"**, execute the full git lifecycle. The workflow is **not complete until the PR is approved and merged**:

1. **Branch** - Create a feature branch from main (`feature/<issue-number>-<description>` or `feature/<description>`)
2. **Commit** - Stage and commit all relevant changes with a conventional commit message
3. **Push** - Push the branch to origin (`git push -u origin <branch>`)
4. **PR** - Create a pull request via `gh pr create` with summary and test plan
5. **Approve PR** - GitHub blocks self-approval on this repo, so post a sign-off comment instead, then merge via `gh pr merge --squash --delete-branch`
6. **Return to main** - `git checkout main && git pull`

## Session Notes

Older session notes (2026-01-19 to 2026-08-31) are archived verbatim in
[docs/session-history.md](docs/session-history.md). Search there for the history
behind a feature or fix; only the current notes are kept here.

### 2026-10-06 - rename-series.ps1 Fixes (branch `fix/rename-series-folder-classification`)

Fixed three #151 bugs, plus -WhatIf:
- **One unit per folder.** A Season folder holding files directly is one table and one
  classification, even when names carry "Disc N" tokens. The FileFilter/ReservedExtras
  per-disc units are gone. Only the play-all check still runs per disc (`-GroupOf` ->
  `Get-SeriesTitleClassification -Groups`). Fixes the lone 4:27 extra becoming an episode.
- **`<Show>-Disc N` folders.** These are now recognised (`Get-DiscFolderNumber`,
  `Get-SuggestedStartEpisode`). Disc folders beside exactly one Season folder join that
  season, and season 0 shares S01's numbering.
- **Numeric sort.** Files are sorted with numbers compared as numbers, so `(2)` comes
  before `(10)`. This also applies to live rips.
- **-WhatIf.** `rename-series.ps1` now accepts -WhatIf (`SupportsShouldProcess`), and it
  always means a dry run.

**Tests:** 472/472 pass (`Test-SeriesRetroRename.ps1` 107).

**Real-file test:** run on the copy `C:\Video\Series\Boys From the Black Stuff`.
- Apply gave the extra to Season 0 and E01-E05 across Season 1, -Disc 2 and -Disc 3.
- Undo fully restored the original names. The copy is left at original names.
- RESOLVED: the 1:42:26 Disc 1 title is "The Black Stuff" (1980 Play for Today) and is now S00-E01 on F: via `-MarkSpecial`.
- RESOLVED: the stale F: manifest was kept as `rename-manifest.undone-old.csv`, and the manifest-append bug is fixed (see item 3 below).

**Added later the same day (also on PR #153), 491/491 tests pass (`Test-SeriesRetroRename.ps1` 126):**
1. **Season 0 (S00).** New Kind `Special`, named `<Title>-S00-E##` in the series-level
   `Season 0` folder. Numbering skips existing S00 files and specials already planned by
   earlier folders in the same run. A title at least 1.6x the expected length (TMDb runtime,
   else the median) is FLAGGED "a special?" but stays an episode, because double episodes
   look the same by length. To make a special: `rename-series.ps1 -MarkSpecial`, `-MarkExtra`
   or `-MarkEpisode` (wildcards on original names), or the prompt's edit option, which takes
   `2s` / `2x` / `2e`. `Get-SuggestedStartEpisode` ignores Special rows.
2. **Explicit `-StartEpisode` wins** for the first folder of each season, even over the
   start suggested from earlier discs.
3. **Manifest-after-undo fixed.** When `undo-rename.ps1` undoes every row it renames the
   manifest to `rename-manifest.undone-<stamp>[-n].csv`, so the next apply starts fresh.
4. **EOF at the confirmation prompt now declines** (it used to accept).

**RESULT on F: (2026-10-06, main 46a3281, `F:\Series\Boys From the Black Stuff`) - done:**
- The stale `Season 1` manifest was kept as `rename-manifest.undone-old.csv`.
- `Season 0\` holds `Boys From the Black Stuff-S00-E01.mp4` (The Black Stuff, 1980 Play for Today, 1:42:26)
  and `-S01-Extra01.mp4` (4:27).
- `-Disc 2` holds S01-E01 and E02. `-Disc 3` holds S01-E04 and E05 (applied with `-StartEpisode 4`).
- Every folder has a manifest and undo script.

**Still open (2026-10-06):**
1. Move the Boys From the Black Stuff `-Disc 2` and `-Disc 3` episodes into `Season 1` (they are still in the `-Disc` sibling folders).
2. Re-rip E03 "Shop Thy Neighbour" (60 min, never ripped) later, then rename with `-StartEpisode 3`. Runtimes: E1 54, E2 57, E3 60, E4 68, E5 68.
3. Set the TMDb key; detection is length-only until then.
4. Blackadder Season 1 has 12 files; files 7-12 are suspected duplicates.
5. Old `Specials` folders (Joking Apart on C: and F:) and legacy `DiscN\extras` folders are not migrated to Season 0 (user: ignore for now).
6. Check that Jellyfin shows Season 0.

### 2026-10-06 - Extras folder renamed `Specials` to `Season 0` (on PR #153)

**Decision (user, 2026-10-06):** `Season 0` is the best name for the series-level extras and
specials folder, so new extras go to `<Series>\Season 0\` (manifest NewName `..\..\Season 0\<name>`
or `..\Season 0\<name>`). The `Specials` name from PR #152 is only read now: `undo-rename.ps1`
still accepts `..\Specials\` rows and legacy `extras\`, and number reservation still reads an
existing `Specials` folder and `DiscN\extras`. "Season 0" matches the season-folder pattern
as season 0, so `Get-SeriesRenameUnits` skips season-0 folders (never a unit to rename).
Existing on-disk `Specials` folders are NOT migrated. The `-Specials` option idea below is unrelated.

### 2026-10-05 - rename-series.ps1 (PR #151, merged 71a74c0)

`rename-series.ps1` + `SeriesRetroRename.ps1` rename already-ripped series folders to
`<Title>-S##-E##` (dry run by default; `-Apply` writes `rename-manifest.csv` and
`undo-rename.ps1` before moving). Handles "Series N" / "<Title>-Series N" folders, "Disc N"
tokens in names, and leaves existing `extras\` alone. 84/84 new tests, full suite passing.
Tested with `-Apply` (apply, undo, re-apply) on the copy `C:\Video\Series\Joking Apart`; not yet on F:.
Known open bugs at the time (ALL FIXED in #153): manifest appended after undo (duplicate rows); EOF at the confirm prompt counts as accept;
mixed Disc-N/no-token files classified per unit; `<Show>-Disc N` folders not matched; no S00/too-long warning;
lexical sort of `(1)..(12)` filenames. PR #152 (extras to Specials) was rebased onto main after #151 merged.

### 2026-10-05 - Series Extras Move to `<Series>\Specials` (PR #152)

**Decision:** Jellyfin did not pick up PR #143's `DiscN\extras\` (Silicon Valley S03); its docs
only recognise extras folders at series or season level. The user chose `<Series>\Season 0\`
beside the Season folders. S##-E## renaming (PR #140) was confirmed working on a real disc.

**What changed (PR #152, `fix/series-extras-specials`, worktree `ripdisc-specials`, 77d3b2f):**
extras are named `<Title>-S##-D#-Extra##[-label]` (disc part avoids collisions in the shared
folder). The manifest stays in `DiscN` with NewName `..\..\Season 0\<name>`. `undo-rename.ps1`
accepts that and legacy `extras\<name>`, and removes Season 0 only when empty. Re-runs reserve
numbers from Season 0 and legacy `DiscN\extras`. 411/411 PowerShell tests pass.

**PR history:** #152 was stacked on #151, then rebased onto main after #151 squash-merged.
The Season 0 layout was exercised on real F: media on 2026-10-06 (Boys From the Black Stuff).

**Next steps:**
1. (Done 2026-10-06: Season 0 verified on F: Boys From the Black Stuff.)
2. Real series rip with extras; confirm they land in Season 0.
3. Check how Jellyfin lists the `-S##-D#-Extra##` files in Season 0.
4. Migrate old `DiscN\extras` folders into Season 0 via rename-series.ps1 (not migrated).
5. Housekeeping: stale worktrees `ripdisc-docs-roadmap` and `ripdisc-core-extraction`
   (branches merged) and a stray `nul` file in the ripdisc root.

### 2026-10-04 - Series Episode Naming, Extras Detection, Rename Manifest and Undo

**What changed:** plain `-Series` Step 3 now renames to `<Title>-S##-E##.ext` (extras
`<Title>-S##-Extra##.ext`), files staying in `DiscN`. This reverses the March 2026
prefix-only change (`a3d4c7c`, `Title-S01-D1-title_t00.mp4`), which had dropped episode
numbers because cross-disc auto-detection raced between concurrent rips. That race is
avoided here by NOT auto-detecting at Step 3: numbering is per disc from `-StartEpisode`,
and Disc 2+ without `-StartEpisode` is asked up front (before the rip), pre-filled from
earlier `DiscN` folders. The answer is carried into the failure-time continue command
(`-StartEpisodeExplicit` on `Get-ContinueRipCommand`).

**Shared code:** the logic lives in `SeriesEpisodes.ps1`, dot-sourced by both scripts
(like `Load-Config.ps1`) instead of being duplicated - a deliberate break from the
copy-in-both-scripts convention, given its size (~500 lines). `undo-rename.ps1` is
standalone and is copied next to each `rename-manifest.csv`.

**Extras detection:** TMDb `/tv/{id}/season/{n}` per-episode runtimes when available
(15% / 3 min tolerance, under 60% = extra), median-length fallback otherwise, play-all =
70-130% of the sum of the others. Durations come from the encoded files (Shell
`System.Media.Duration`, HandBrakeCLI `--scan` fallback), not MakeMKV `TINFO:n,9` - the
disc is already ejected by Step 3 and `TINFO` is only parsed in auto-discovery mode.

**Testing status:** `tests/Test-SeriesEpisodeRename.ps1` 88/88; full suite 237/237.
Duration reader checked against one real MP4 (7104 s); MKV duration via the Shell
property is unverified. **Not exercised against a real rip or the live TMDb API**
(TMDb is mocked in the tests).

**Outstanding:** real-disc validation of a Disc 1 + Disc 2 season; check whether Jellyfin
treats `-S02-Extra01` files as extras or ignores them (an `extras\` subfolder is the
alternative); the C# `-processQueue` path does not get the new naming.

### 2026-10-04 (continued) - TheDiscDB Lookup for Series Episode/Extra Mapping

**What changed:** plain `-Series` naming now asks TheDiscDB (thediscdb.com) first, then TMDb
runtimes, then the median heuristic. Code is in `SeriesEpisodes.ps1` (TheDiscDB section at the
end); `-NoDiscDb` on both scripts, `-DiscDbHash` on `continue-rip.ps1`.

**API facts (researched 2026-10-04):** public GraphQL at `https://thediscdb.com/graphql`, no key,
HotChocolate filtering (`where:` on `mediaItems`, `releases`, `discs`). A disc is identified by
`contentHash` = uppercase hex MD5 over each file's size as little-endian Int64, files ordered by
name: Blu-ray = `BDMV/STREAM/*.m2ts` (direct children), DVD = every file directly in `VIDEO_TS`
(source: TheDiscDb/web `DiscScanner.cs`, `HashingExtensions.cs`). Each `Title` has MakeMKV's
`index`, `sourceFile`, `duration` (h:mm:ss), `size`, and `item { title type season episode }`
(`type` = Episode / DeletedScene / Extra / Trailer / ...; no item = unidentified). Response is
`charset=utf-8`. Also available but unused: `globalDiscId` (AACS/DVD disc id) and `fingerprint`.

**Why matching is not by `_tNN` index:** TheDiscDB lists every MakeMKV title (even 9 s clips), so a
user's MakeMKV minimum-length setting shifts `_tNN`; Step 2 also skips the composite. Durations
alone are ambiguous (30 Rock S1D2 has two 0:21:35 episodes). `Get-DiscDbTitleMatches` does an
order-preserving DP alignment on duration (5 s / 1% tolerance), most pairs first, then closest
durations, then least index drift.

**Hash timing:** computed before the rip from the drive listing (disc is ejected after Step 1).
Skipped with `-DriveIndex` but no `-Drive`, because `$driveLetter` is then only the config default.

**Testing status:** `tests/Test-TheDiscDbLookup.ps1` 78/78; full suite 315/315. Live API checked
by hand with 30 Rock S1D2's hash (match, 2 s round trip, ~3 KB). Real-drive check: a Silicon
Valley S2 D2 DVD in E: hashed in 267 ms (27 VIDEO_TS files) and fell back cleanly - but TheDiscDB
has no Silicon Valley entries at all, so a real-disc hash MATCH is still unconfirmed (only the
independent Python implementation agrees). TheDiscDB coverage is mostly Blu-ray; expect DVD misses.
**Not exercised in a real rip.**

### 2026-10-04 (continued) - C# Core Extraction (step 1 towards a WinForms app)

**Decision (user, this session):** build the WinForms app as a new project inside this repo
(`RipDisc.WinForms`, not yet created), on a shared `RipDisc.Core` library, rather than in a new
repo or by wrapping the PowerShell. Still open: whether the PowerShell scripts get frozen once the
C# side catches up, or both stay maintained (every fix twice).

**What changed:** `RipDisc/` is now `RipDisc.sln` with `RipDisc.Core`, `RipDisc.Cli` (still
`RipDisc.exe`) and `RipDisc.Tests` (xUnit, 62 tests - the first C# tests here). The pipeline
(`RipPipeline`, ex-`RipDiscApplication`) talks only to `IRipUI`; `IProcessRunner` and
`IRipEnvironment` (eject, open folder, drive-ready check, handle wait) are seams so tests run the
whole pipeline against fakes in a temp folder. Config now comes from `ripdisc-config.json`.
Details in CHANGELOG.

**Found and fixed:** `-processQueue` merged the just-finished job back in from the queue file
(see CHANGELOG). Long-standing, only in the C# path.

**Worktree:** built in `C:\Users\sjbeale\source\repos\ripdisc-core-extraction` because another
session switched the main checkout to `fix/series-extras-subfolder` mid-task.

**Testing status:** 62/62 tests; CLI smoke-tested (bad-argument usage, closed-stdin abort).
**Not run against a disc.**

**Next for the WinForms track:**
1. `makemkvcon -r` robot-mode spike on a spare disc (progress `PRGV`/`PRGC`/`PRGT`, `MSG` codes,
   licence-expiry text) - feeds a progress bar and better error analysis
2. `RipDisc.WinForms` project: a `WinFormsRipUI` that marshals `IRipUI` calls to the UI thread,
   running `RipPipeline.Run` on a background task with a Cancel button wired to the token
3. Port PowerShell-only behaviour the GUI needs most (see README Feature Parity table)

### 2026-10-04 (continued) - MakeMKV Robot-Mode (-r) Trial, PAUSED

**Status:** started, then paused by the user. Deliberately NOT run against a drive: two live rips
were in progress (Silicon Valley S2 Disc2 on disc:2, S4 Disc1 on disc:0), and an info query makes
MakeMKV scan every drive, which has interfered with concurrent rips before. Only static analysis
and an existing captured fixture were used.

**Existing fixture:** `%TEMP%\makemkv-drive-cache.txt` (real `-r` output, MakeMKV v1.18.4).
Contains an `MSG:1005` start line, `DRV:index,visible,enabled,flags,"drive name","disc name","D:"`
lines (e.g. `DRV:0,2,999,1,"BD-RE HL-DT-ST BD-RE BU40N 1.05 MO4P6N95940","SILICON VALLEY S1 D2","D:"`;
empty slots are `DRV:n,256,999,0,"","",""`), `MSG:5010` "Failed to open disc", `TCOUNT:0`.
Next session: copy it into `RipDisc.Tests` as a fixture.

**Robot format strings (from makemkvcon64.exe):** `MSG:%u,%u,%u,"..."`, `PRGV:%u,%u,%u`,
`PRGT:%u,%u,"`, `PRGC:%u,%u,"`, `DRV:%u,%u,%u,%u,"`, `TCOUNT:%u`, `CINFO:%u,%u,"`,
`TINFO:%u,%u,%u,"`, `SINFO:%u,%u,%u,%u,"`.

**Progress/operation strings:** "Current progress - %u%%  , Total progress - %u%%", "Current
operation: %s", "Current action: %s", "Saving %1 titles into directory %2", "Copy complete. %1
titles saved, %2 failed.", "Operation successfully completed", "Failed to save title %1 to file
%2"; operation names include "Opening DVD disc", "Opening Blu-ray disc", "Processing BD+ code...",
"Decrypting data".

**Licence strings (for a "MakeMKV key expired" detector):** "Evaluation period has expired,
shareware functionality unavailable.", "Evaluation version, evaluation period expired %1 day(s)
ago", "Your temporary key has expired and was removed. Please restart the application.", "This
application version is too old.  Please download the latest version at %1 or enter a registration
key...", "The stored activation key is invalid...".

**KEY FINDING:** "You are trying to start MakeMKV evaluation from a third-party application.
Please launch MakeMKV if you would like to start the evaluation period." A CLI/GUI wrapper cannot
start the Blu-ray evaluation; the MakeMKV GUI must be opened once. Answers part of the earlier
licensing question.

**Unknown:** numeric MSG codes other than 1005 and 5010 - need a real `-r` run.

**Remaining trial steps (only when no rips are active):**
1. `makemkvcon64 -r --progress=-same info disc:N` on a spare disc; capture full output as a fixture
2. C# `MakeMkvRobotParser` in `RipDisc.Core` with tests
3. `RipDisc.WinForms`: `WinFormsRipUI` marshalling `IRipUI` to the UI thread, pipeline on a
   background task, Cancel wired to the `CancellationToken`, progress bar from `PRGV`

**PR #144 test plan still unchecked:** real disc rip with `RipDisc.exe`; `-processQueue` with two
real jobs. Merged to main 2026-10-04 without a real-disc test.

### 2026-10-04 (continued again) - Series Extras Subfolder; `continue-rip.ps1 -Yes` for Series Naming

**Extras location:** plain `-Series` extras now move to `DiscN\extras\` (`$script:SeriesExtrasFolder`
in `SeriesEpisodes.ps1`), matching movie mode's lowercase `Join-Path $finalOutputDir "extras"`.
Per disc, not per season: Extra## is numbered per disc (season level would collide), concurrent
disc rips must not share a folder (PR #41's reason for DiscN), and manifest/undo stay per folder.
Plan rows: `NewName` = path relative to DiscN (`extras\<file>`), `FileName` = bare name. Renames
use `Move-Item` (no `-Force`). Manifest columns deliberately unchanged - adding one would break
`Export-Csv -Append` onto PR #140-era manifests.

**Undo:** `undo-rename.ps1` whitelists exactly `^(?:extras[\\/])?[^\\/:]+$` for `NewName` (plus no
`.`/`..`), moves files back with `Move-Item`, removes `extras\` only when empty.

**`-Yes`:** `continue-rip.ps1 -Yes` -> `Get-AutoStartEpisode` (suggestion, else fallback flagged as a
guess) and `Invoke-SeriesEpisodeRename -AutoAccept`; both logged. `rip-disc.ps1` has no `-Yes` or
other non-interactive flag, so it was not changed (adding one would also have to cover the
Ready-to-rip, title-warning and drive prompts - out of scope).

**Known limitation (deferred by the user):** the C# `-Queue`/`-processQueue` path has none of the
PR #140/#142/this-PR series naming; it writes series files into the Season folder with the older
`Title-S##-` prefix (`PrefixSeriesFiles`). Documented in README.

**Testing status:** `Test-SeriesEpisodeRename.ps1` 109/109 (21 new), `Test-TheDiscDbLookup.ps1` 79/79.
**Not exercised in a real rip**; Jellyfin's handling of `DiscN\extras\` (an extras folder nested
under a non-season folder) is unverified.

### 2026-10-04 - MakeMKV Progress Bar, Milestones and ETA (PR #141)

**What changed:** `rip-disc.ps1` now runs MakeMKV with `--progress=-same` and shows a
`Write-Progress` bar, 10% milestone lines with ETA, and op/action lines. Progress lines are
neutral to the stuck-sector watchdog (they neither reset nor trip it), and the rip-started
marker is narrowed to `Saving N titles|Title #`. Format strings were verified against
`makemkvcon64.exe` v1.18.4. New tests: `tests/Test-MakeMkvProgress.ps1` (9/9 pass).

**Status:** PR #141 squash-merged to main at f52eea0 on 2026-10-04 without a real-disc test; **not hardware-tested** - validate the bar/ETA and watchdog on the next rip.
Developed in the worktree `ripdisc-makemkv-progress` (sibling of the main checkout) because
the main checkout held another session's uncommitted `SeriesEpisodes.ps1` changes on
`feature/thediscdb-lookup`.

**Outstanding:** rip a real disc from this branch and confirm the bar/ETA and watchdog
behaviour. Rebased onto `main` after #142/#143/#145 (TheDiscDB/extras work) landed:
`rip-disc.ps1` merged cleanly; only CHANGELOG/CLAUDE.md needed both entries kept.

### 2026-10-05 (continued)

**C: test of #151** (Joking Apart copy, `C:\Video\Series\Joking Apart`): -Apply, undo and re-apply all worked. The copy was left renamed; D1_t00 was marked as an extra by hand.

**Fix committed on `feature/series-rename-utility`:** `SeriesRetroRename.ps1` (~line 97) FileFilter closures used `.GetNewClosure()`, which cannot see `Get-DiscNumberFromFileName` (dot-sourced into rename-series.ps1's script scope), so folders with "Disc N" tokens failed with "not recognized". Now captures `$discOf = ${function:Get-DiscNumberFromFileName}` and calls `& $discOf $n`. Verified on temp folders; 11 test files pass.
**(Still true:) the new regression test in `tests/Test-SeriesRetroRename.ps1` (rename-series.ps1 via `powershell.exe -File` on the Joking Apart fixture) is INEFFECTIVE:** it still passes against the old buggy code. Marked with a TODO; needs reworking so it fails on the old code.

**PR #152** was rebased and merged (extras folder, now named `Season 0`).

**Open bugs (a-e ALL FIXED in #153, 2026-10-06; kept as history):**
- (a) A folder mixing files with and without "Disc N" tokens is split into separate units: one prompt per unit, episode-vs-extra judged per unit, and a lone file is always the unit median so always becomes an Episode. Real case: `F:\Series\Boys From the Black Stuff\Season 1` - B1_T00-1.mp4 (4:27, 47 MB, clearly an extra) became S01-E01 and the 1:42:26 Disc 1 title became S01-E02. A normal rip-disc -Series run judges all of a disc's titles together. Fix: one table, one prompt, one classification per folder.
- (b) Disc folders named `<Show>-Disc N` not recognised: `^disc\s*(\d+)$` at SeriesRetroRename.ps1:29 needs a prefix allowance like the season regex. Boys From the Black Stuff-Disc 2 and Disc 3 were never processed.
- (c) No specials/S00 concept: only too-SHORT titles are flagged, not too-long. The 1:42:26 title may be "The Black Stuff" (1980 play), i.e. a special. Ideas: `-Specials` parameter; table warning for titles much longer than the median.
- (d) After undo, manifest and undo script remain and the next apply appends to the manifest (duplicate rows).
- (e) End of input at the confirmation prompt counts as "accept" (SeriesEpisodes.ps1 ~541).

**F: status (RESOLVED 2026-10-06):** the stale Season 1 manifest was kept as `rename-manifest.undone-old.csv` and Boys From the Black Stuff was re-run correctly (see the RESULT above).

**Unfinished:** investigate how ripdisc's -Series rip logs detect duration discrepancies, using the Silicon Valley runs as the example (logs: `C:\Video\logs\Silicon Valley_disc*_*.log`). Logs found cover Seasons 1-5; none for Season 6 yet. The user says the runs they meant may be from an EARLIER season, so check the Season 1-5 logs. Then model the retro-rename classification on what the live rip does.
