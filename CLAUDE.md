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

**Status:** PR #141 deliberately left OPEN pending a real-disc test; **not hardware-tested**.
Developed in the worktree `ripdisc-makemkv-progress` (sibling of the main checkout) because
the main checkout held another session's uncommitted `SeriesEpisodes.ps1` changes on
`feature/thediscdb-lookup`.

**Outstanding:** rip a real disc from this branch and confirm the bar/ETA and watchdog
behaviour. Rebased onto `main` after #142/#143/#145 (TheDiscDB/extras work) landed:
`rip-disc.ps1` merged cleanly; only CHANGELOG/CLAUDE.md needed both entries kept.
