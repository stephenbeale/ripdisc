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

### 2026-11-01 (evening) - Year-match order, title sanitising, manifest append (PRs #182, #183, #184 all OPEN)

**Open PRs (none merged this session; real-rip checks owed for all):**
- #182 `fix/year-match-word-order` (4b6759b, main checkout): `Test-SimilarTitle` needs shared words in the same order (typed "The Story of Concorde" was offered "Concorde: The Untold Story" (2023)). Year prompt [N] "Never ask again" saves `askFolderYear=false` via new `Save-ConfigSetting` (Load-Config.ps1); setup.ps1 step 7 asks/keeps it; sample config `askFolderYear:true`. 1168 tests / 23 files pass.
- #183 `fix/title-sanitize-no-underscore` (53314fa, worktree `ripdisc-sanitize`): `Get-SafeTitle` (rip-disc.ps1 + continue-rip.ps1) removes `\ / : * ? " < > | [ ]` (run becomes a space, spaces collapsed) instead of `_`; `Get-LegacySafeTitle`/`Resolve-LegacySafeTitle` keep existing `_` folders; IMDb kind "video" shown as [Direct-to-video]. The live Strike Force rip ran on old code, so `F:\Documentaries\Strike Force_ The Story of Sukhoi (2007)` exists with `_`; rename offered, user has NOT decided - do not rename.
- #184 `fix/manifest-append-columns` (570b924, worktree `ripdisc-manifest`): `Write-RenameManifest` appends to manifests with extra columns; if the manifest cannot be written nothing is renamed. Trigger: rename-series on `F:\Series\Silicon Valley\Season 3\Disc1` at 18:26 renamed "Silicon Valley - S03E0N - <title>.mp4" to "Silicon Valley-S03-E0N.mp4" while the manifest append failed. The 5 rows were hand-added (backup `rename-manifest.backup-20261101.csv` beside it). Episode titles lost from names; user undecided whether to restore. Do NOT run undo-rename.ps1 there (it would undo ALL rows back to raw A1_t00 names).

**Findings:**
- Slow start: rip-disc.ps1's own pre-rip work is ~12 s; MakeMKV "Scanning CD-ROM devices" took 41 s because the BU40N and the second GP75N dropped off USB (PnP Present=False). Hardware, not code.
- Mistake: a scripted test answered "A" at "Press Enter to continue" (does not abort) and started a 2nd MakeMKV on D: during the "Super Carrier Airport at Sea" rip, ~17:11:33-17:12:20 (killed). Check that rip's log for read errors around then. Never script rip-disc.ps1 past the Ready-to-rip prompt.
- Set `CLAUDE_CODE_USE_POWERSHELL_TOOL=1` in `C:\Users\sjbeale\.claude\settings.json` env.
- Untracked `ripdisc-config.json.bak-20261029` and git-ignored `ripdisc-config.json` in the main checkout: leave alone.

**Ideas not done:** "Pick a title" prompt should accept typed text as the title (user typed one, got "Invalid choice"); warn before MakeMKV when disconnected drives will slow the scan; rename-series should leave already-titled S##E## files alone.

**NEXT SESSION:**
1. Real-rip checks for #182 (year prompt order + never-ask), then review/merge #182, #183, #184 (stacked-PR merge gotcha: check bases).
2. Ask user: rename `Strike Force_ The Story of Sukhoi (2007)` to the new no-underscore name? Restore Silicon Valley S03 episode titles?
3. Check the Super Carrier Airport at Sea rip log for read errors ~17:11-17:12.
4. Fix the USB drop-offs (BU40N, GP75N) - see WD/USB hardware memory.
5. Implement the "Pick a title" free-text option and the disconnected-drive warning.

### 2026-11-01 - Library prompt, setup library folders, year prompt similarity (branch `feature/library-picker`, worktree `ripdisc-library`)

Trigger: real Duxford DVD run. The user wanted to pick a library (Documentaries, Music ...) after the output drive, setup.ps1 to ask for library folders and suggest them, and the year prompt (after editing the title to "Duxford the Second World War Years") offered six unrelated films - skip it when nothing is similar.

**Built:** `Read-LibraryChoice` / `ConvertTo-LibraryChoice` / `Test-SimilarTitle` (Load-Config.ps1); `MediaLibraries.ps1` (scan + "likely library" heuristic, shared with setup.ps1); setup step 5/7 saves `libraryFolders`; `Read-FolderYear` filters by similarity and `-SkipWhenNoSimilar` (the main year prompt only, not the clash fallback). Library prompt only runs when the output drive was prompted and no type switch was given. Also fixed `$name?` in the -Basketball prompt.

**Checked live (worktree):** setup.ps1 with scripted answers on the real F: (ticks correct, config written to the worktree's own git-ignored config); real rip-disc.ps1 via a fake Read-Host: drive F -> library list -> Documentaries -> discovery (D: empty). Full suite 1144/1144 across 23 files. **Not run on a real rip**; user's main-checkout config has no `libraryFolders` yet (run setup.ps1 or the prompt scans F: each run).

### 2026-10-29 (late) - Rip stats, auto drive pick (PRs #170, #171) and NEXT SESSION items (main = 35606dc, all 19 test files pass)

Both PRs MERGED, **neither run on a real disc**.

**PR #170 (squash e93234b):** new `rip-stats.ps1` + shared `RipStats.ps1` (per-type rip time stats and outlier reasons from `C:\Video\logs` and `C:\Music\logs` plus each album's cyanrip log; `-JsonPath` export). `rip-disc.ps1` logs "Estimated finish: ~HH:mm (median of N past <Type> rips ...)". The "Ready to rip" summary always shows the full "Output folder:" path incl. category. Fix: a title already ending "(YYYY)" got the year twice (`Get-MediaFolderName` replaces a trailing year, new `Get-TitleYear` skips the year prompt). Tests: Test-RipStats 29/29, Test-MovieFolderNaming 125/125.

**PR #171 (squash 35606dc):** without `-Drive`, if the default drive holds no DVD/Blu-ray (audio CD, empty, not in MakeMKV's list) or is busy, switch to the one other free drive with a video disc, or ask if several (`Select-AutoVideoDrive`); when nothing better, name PnP-not-present optical drives (`Get-DisconnectedOpticalDrives`). Test-AutoVideoDrive 21/21.

**Findings from rip-stats on real logs:** median start-to-complete: DVD movie 51 min, documentary 52 min, TV series disc 1h07m, album 30 min. About 20% cyanrip paranoia VERIFY is baseline; >1000 fix-ups = damaged disc. Every no-damage slow album was on one of the two HL-DT-ST DVDRAM GP75N drives (median 0.6x real time vs BU40N 0.47x). Logs record no disc region. Dashboard artifact (snapshot): https://claude.ai/artifact/YXUE9569xWopJPHEdZrzNf

### 2026-10-29 (late) - NEXT SESSION items

1. Real-disc checks for #170: Estimated finish line, Output folder line, no doubled year and no year prompt for a "Title (YYYY)" pick.
2. Real-disc checks for #171: audio CD in the default drive + DVD elsewhere switches; two DVDs asks; the disconnected-drives line.
3. Rename doubled-year folders/files once the encodes finish: `F:\DVDs\Jean de Florette (1986) (1986)` and `F:\Surf\Musica Surfica (2008) (2008)` (made before the fix). Ask the user first.
4. Hardware: at 18:5x on 2026-10-29 the BU40N and the second GP75N (fw 1.02) showed Present=False in PnP (dropped off USB); only GP75N 1.01 (D:, audio CD in it) was visible. Reconnect/check the hub (see also the WD Elements/USB root-hub memory).
5. Audio CD still in D: - rip it with ripaudio.
6. Empty folder `C:\Users\sjbeale\source\repos\ripdisc-auto-drive` left behind (was locked by the session) - delete it.
7. Idea not done: ripaudio could warn when a GP75N is picked for audio (slower reads).
8. Still open from the earlier lists below (not repeated): the 2026-10-29 items 1-6 (Step 3 image deletion, continue-rip.ps1 clash/ForeignFeature guard, C# RipDisc.Core lacks #166 fixes, #166 real-disc checks, re-rip 1983 Scarface) and the 2026-10-28 items 1-6 (cable-pull test for #161, README CR lines, Write-Host spacing, real-drive tests, Slow Dance IMDb, ripaudio/ripdisc direction). Note PR #168 has since fixed the image-deletion and continue-rip risks per memory; verify before acting. The untracked `ripdisc-config.json.bak-20261029` is still undecided: do not commit, delete or move.

### 2026-10-29 - Ask for output drive, custom genre folders, offline lookups (PR #169, merged 78c6c6e; 911/911 tests, 17 files)

Trigger: `.\rip-disc.ps1 -Drive G` (no -OutputDrive) sent Jean de Florette to E: silently, via config defaultOutputDrive "E:". The user also wanted unknown switches like -Basketball treated as new folders, and to be told "cannot check title" and carry on when offline. And the "Discovered Metadata" block said `Type: Movie` for a -Documentary run (only the label was wrong; the path was already Documentaries).

**Built:** `Read-OutputDrive` prompt (Enter = config default; the answer is treated as explicit for the continue-rip / ripaudio commands and the [D] relaunch). `-Genre <name>` on both scripts, plus `ValueFromRemainingArguments $UnknownArgs`. `Resolve-CustomGenreArg` asks Y(Enter)/N/A: N = standard DVDs/Bluray folder, and a near-miss typo is flagged via `Get-EditDistance`. Other unknown args exit 1; -Genre with a built-in genre exits 1; a bad folder name exits 1. `Test-InternetConnection` (HTTPS probe, 3 s, live 310 ms) runs once via `Test-OnlineOnce`; offline it skips TheDiscDB/TMDb/IMDb, adds `-Offline` on `Get-SeriesDiscDbLookup` (the hash is still read), drops [R], and `Get-YearLookup` returns $null. `Get-ContentTypeLabel` replaces every inline content-type chain. The helpers live in Load-Config.ps1.

**Tests:** tests/Test-OutputDriveAndGenre.ps1 52/52 (also runs the real rip-disc.ps1 early-exit paths). Test-MovieFolderNaming imports Test-OnlineOnce/Get-YearLookup. C# RipDisc.Core unchanged.

**Still owed (PR #169 test plan):** a real rip with -Basketball, a real run without -OutputDrive, a real offline run.

### 2026-10-29 - First REAL mid-rip drive drop (PR #161 behaviour)

- Jean de Florette DVD on G:. The drive dropped at ~70% of the MKV save (~18:28; STATUS_DEVICE_NOT_CONNECTED). PAUSED prompt -> R -> re-query found "G: is now disc:1 (was disc:0)" and the re-rip started. Verified live: the prompt, retry, and the renumbered-drive lookup.
- Not yet confirmed: that the retry completes, and that the partial MKV from attempt 1 was removed (check C:\Video\Jean de Florette (1986)\Disc1).
- Cost: ~18 min of work redone, ~26 min later finish (expected ~19:00 instead of ~18:34); the retry was slightly slower.
- New cosmetic defect: attempt 1's Write-Progress bar ("Total 21% | Elapsed 6m 18s") stayed on screen into attempt 2. Needs `Write-Progress -Completed` / a reset between attempts.
- Possibly the same USB instability as the failing F: drive (WD Elements); a wiggle-test script is ready but not run.

### 2026-10-29 - NEXT SESSION items (added after PR #169)

1. Three real checks for PR #169: a real rip with -Basketball, a real run without -OutputDrive, a real offline run.
2. Confirm the Jean de Florette retry finished and the partial MKV from attempt 1 was removed (C:\Video\Jean de Florette (1986)\Disc1).
3. Fix the stale Write-Progress bar carried from attempt 1 into attempt 2 (`Write-Progress -Completed` / reset between attempts).

### 2026-10-29 - Scarface folder clash fix (PR #166, merged 7ff567a) and NEXT SESSION items (main = 7ff567a, 809/809 tests, 16 files)

Trigger: the 2026-10-28 real rip of Scarface (1932), picked from the pick list (title "Scarface", year 1932), went into `F:\DVDs\Scarface` (the 1983 film's folder). Config used plain `title` folder naming and the existing-same-name-folder rule beat the year, so Step 3 saw "Feature already exists" and moved the new film into extras. Log: `C:\Video\logs\Scarface_disc1_20261028_202002.log`. The user later moved the 1932 film to `F:\DVDs\Scarface (1932)\Scarface (1932)-Feature.mp4` (1.89 GB, fine).

**The 1983 file was NOT damaged by that run:** `F:\DVDs\Scarface\Scarface-Feature.mp4` is 55 MB, has no moov atom (unplayable), is dated 13/01/2026, and the 10-28 log listed it at 0.05 GB before anything moved. It has been broken since January and needs a re-rip into `F:\DVDs\Scarface (1983)`.

**PR #166 (squash 7ff567a):** pre-rip folder clash check for Disc 1 / documentary part (switches to `Title (Year)` via the picked year or the year prompt, else stops before ripping); the picked year's existing `Title (Year)` folder is preferred over a plain same-name folder; HandBrake outputs de-clashed with `Get-UniqueFilePath` (no silent `-o` overwrite; the recovery script uses the same names); Step 3 ForeignFeature guard (never sweeps into extras beside an existing Feature). New `Get-ExistingFeatureFile` in Load-Config.ps1. Test-MovieFolderNaming 69/69 (12 new tests fail on old main); full suite 809/809. **Not run on a real disc.** continue-rip.ps1 deliberately left with BaseName.mp4 encode naming (existing MP4 = already encoded, for resume).

**Config:** the user's local git-ignored `ripdisc-config.json` now has `movieFolderNaming = "title-year"`. Backup is the untracked `ripdisc-config.json.bak-20261029` (user has not decided; do not commit, delete or move).

**Audit answer** ("will any titles overwrite a main feature from now on?"): no code path overwrites a `-Feature` file now (Rename-Item/Move-Item never use -Force; encode outputs are raw MakeMKV names, now made unique). Remaining risks are items 1-3 below.

### 2026-10-29 - NEXT SESSION items

1. **Step 3 deletes every jpg/png/gif/bmp in the output folder** (rip-disc.ps1 ~line 3300, continue-rip.ps1 ~line 1499). A Disc 2 / `-Extras` / continue run into an existing film folder wipes that film's Jellyfin artwork (folder.jpg, backdrop.jpg etc.); probably why F:\DVDs\Scarface's artwork was rewritten at 21:04 on 10-28. Fix idea: only delete images not present before the run, or skip deletion when the folder pre-existed.
2. **continue-rip.ps1 has no pre-rip folder clash check and no ForeignFeature guard**, so a continue run can still sweep a new film into extras beside an existing Feature. Its `-Force` re-encodes over its own raw names only.
3. **C# RipDisc.Core has none of #166's fixes:** RipPipeline.cs:546 builds the HandBrake output as BaseName.mp4 with no uniqueness, and RipPipeline.cs:574 deletes outputFile on Blu-ray subtitle retry, so a pre-existing raw-named mp4 (not a Feature) could be overwritten or deleted.
4. Real-disc checks for #166: a film whose plain folder holds a Feature gets `Title (Year)`; the refusal message when there is no year.
5. Re-rip the 1983 Scarface into `F:\DVDs\Scarface (1983)`; delete the broken 55 MB `F:\DVDs\Scarface\Scarface-Feature.mp4` only with user confirmation.
6. Still open from the 2026-10-28 list below: cable-pull test for #161, README CR lines, Write-Host spacing, other real-drive tests, Slow Dance IMDb question, ripaudio/ripdisc direction.
7. Friend's release safety: undecided, (a) tags / (b) stable branch / (c) README rollback section; bad release = git revert on main, never force-push.

### 2026-10-28 - NEXT SESSION items (main = 9f1b240, 790/790 tests, 16 files)

All PRs below are MERGED: #162 e362198, #164 c919c5b (replaced #163, auto-closed when its stacked base branch was deleted), #156 8621a43, #161 dbbb24c, #159 9f1b240. #156, #161, #159 were merged without real-hardware tests.

1. **Cable-pull test for #161 is still owed on main** (junk disc, pull USB mid-rip, check beep + tray notification + PAUSED prompt, reconnect, R, rip completes, partial MKV removed).
2. Fix README.md (from #159): two example lines contain a literal carriage return in the middle of the rip-disc.ps1 path (the backslash-r of `.\rip-disc.ps1` was turned into a CR).
3. Fix #159's `Write-Host "Disc: ..."` spacing defect in rip-disc.ps1 (no space before `-ForegroundColor`).
4. Real-drive tests not yet done: audio-CD stop prompt and `[D]` restart; SCARFACE pick-list flow; setup step 6 and the year prompt; Disc 2 reusing a dated folder; #156 watch-drives with a real insert; #159 multi-disc documentary rip.
5. Slow Dance DVD is a 2014 surf film; IMDb tt3453042 (a short) is NOT confirmed as the same film.
6. ripaudio/ripdisc direction (project-manager advice, user has not chosen): backlog cleared; next build a front-door `rip.ps1` launcher in ripdisc dispatching on disc type, later a shared RipCommon module (drives, disconnect/retry, logging, config). Do NOT fully merge ripaudio into ripdisc. ripaudio has no tests. Proposed for ripaudio, not started: text-search pick list when MusicBrainz has no disc-ID match (rip-audio.ps1 ~line 1373), richer multi-release picker (~line 402), -album vs disc sanity check.

### 2026-10-28 - Movie folder naming `Title (Year)` (PR #164, merged c919c5b)

User chose option 3: only `Title` / `Title (Year)`, no IMDb id yet. Setup step 6 saves `movieFolderNaming` (`title` / `title-year`); `-Year` on both scripts. Precedence: `-Year`, then exactly one existing `<title>`/`<title> (YYYY)` folder, then (title-year) the discovery pick's year or the `Read-FolderYear` prompt (Enter/EOF = no year, never aborts). Only the 7 movie/genre-film `$finalOutputDir` lines use `$script:MediaFolderName`; series untouched. Tests: Test-MovieFolderNaming 50/50, full suite 670/670 (13 files). Load-Config read checked against real config files (title-year / bogus / none). **Not run on a real disc.**

**Deferred:** `[imdbid-tt...]` naming needs a bracket audit first (wildcards in `-like` and `-Path`: Get-Item/Test-Path/New-Item/Get-ChildItem on `$finalOutputDir`, the movie-mode prefix `-like` checks in both scripts).

### 2026-10-28 - Title suggestions from the disc (PR #162, merged e362198)

Trigger: user ran `rip-disc.ps1 -title "Slow Dance" -Surf -Drive d -OutputDrive F` while the drive listing already showed `[Slow Dance]`. They asked for the title to be inferred from the disc, with API suggestions, years and IMDb ids.

**Built:** without `-title`, auto-discovery now shows one numbered pick list from TheDiscDB (hash, no key), the cleaned disc label, TMDb (key, plus `external_ids` for IMDb ids) and IMDb's suggestion feed `v2.sg.media-imdb.com/suggestion/<c>/<q>.json` (no key, UNOFFICIAL). Rows show year, `[Movie]`/`[TV]` and IMDb id. With `-title`, the label from the drive listing is compared with the title (NOTE on mismatch, no prompt). The discovery TheDiscDB lookup is reused by `-Series` naming. `tests/Test-TitleSuggestions.ps1` 79/79; full suite 620/620 across 12 files.

**Real-disc run (user, Slow Dance DVD):** the menu worked but the user says the film is 2014 and no 2014 row appeared. IMDb's suggestion feed returns only ~8 hits ranked by popularity, and "slow dance 2014" does find `tt3453042` (a 2014 short; NOT confirmed to be the user's surf film). Added a description line (TMDb overview / IMDb cast) and a check link per row, IMDb's real kind, and `[R] Search again`.

**Second real-disc run:** an audio CD in D: (label "Audio CD") was offered film titles. Fixed: an AUDIO CD GUARD right after the drive match (reuses `Get-DiscTypeLabel`, checked live: D: -> Audio CD, G: -> DVD-Video) prints the ripaudio command and offers [R] start rip-audio.ps1 / [D] rip another drive's video disc. Also `[D] Use a different drive` in the pick list (restarts the script via `Get-RelaunchParams` splat - verified with a throwaway script that switches and exit codes survive), and `[Movie]` hidden (user request). Test-TitleSuggestions 99/99 at that point; the guard and [D] restart have NOT been run end-to-end on a real drive.

**Found live and fixed before commit:** a year-less label merged into IMDb's first same-name hit (2005 Japanese TV series "Slow Dance"), so Enter would have turned on Series mode for a surf film. The bare label now never absorbs search hits; only TheDiscDB does.

**Not done:** no real-disc run of the pick list. TMDb key still unset, so the TMDb path is mocked only. IMDb endpoint verified live 2026-10-28 (about 0.4 s; one 8 s cold-start timeout seen once). Year/IMDb id are shown and logged but NOT used in folder/file names (Jellyfin `Title (Year) [imdbid-tt...]` naming is a possible follow-up). C# side unchanged.

**Still open from before:** PR #161 cable-pull test (#161 now merged; test still owed on main).

### 2026-10-26 - Pause-and-retry on mid-rip drive disconnect (PR #161, MERGED dbbb24c)

**NEXT SESSION: do the cable-pull test for PR #161 on main (merged without it at the user's choice). The user asked to be reminded.**

1. `git checkout main && git pull`.
2. Start a rip on a junk disc.
3. Pull the drive's USB cable partway through.
4. Confirm the beep, the tray notification and the PAUSED prompt.
5. Reconnect, press R, and confirm the rip completes and the partial MKV was removed.
6. Tick the last test-plan box on PR #161.

Trigger: a Cream documentary rip failed when the user knocked the USB cable (MakeMKV `STATUS_DEVICE_NOT_CONNECTED`; Step 1 failed "The drive disconnected ... while MakeMKV was saving titles"). Nothing completed, so nothing is resumable.

**Built (9aa7690, branch off main 87293d4, PowerShell only, C# RipDisc.Core unchanged):** rip-disc.ps1 Step 1 beeps and shows a tray notification (user confirmed it appears), removes that attempt's MKVs, asks [R]etry/[A]bort. R is accepted only once Windows sees a readable disc. Re-queries MakeMKV's disc:N with the drive cache cleared, re-rips the whole disc, max 3 attempts. New tests/Test-DriveDisconnectRetry.ps1 27/27; full suite 568 tests across 12 files. Loop exercised against a fake makemkvcon (renumbered drive, abort, 3 disconnects in a row, -DriveIndex mode). **Untested: a real drive unplug.**

**Also open:**
- Re-rip Cream (Documentary); it still needs doing.
- Possible follow-ups: retry only unfinished titles; C# parity.

### 2026-10-21 - rename-series handles `S1` folders, file-name episode numbers, `._` files (PR #157, merged 45369cf)

Trigger: `rename-series.ps1` on `F:\Series\Parks and Recreation` found only 2 of 6 season folders (`Parks and Recreation S1/S2/S3/S6`, `Season 4`, `season 5`) and macOS `._` AppleDouble files were numbered as episodes ahead of the real files.

**Fixed in #157 (541/541 tests: Test-SeriesEpisodeRename 162, Test-SeriesRetroRename 135):**
- `Get-SeasonFolderNumber` (SeriesRetroRename.ps1) and `Get-SeriesSpecialsRelativeDir` (SeriesEpisodes.ps1) accept `<Title> S1`, `S01` and bare `S6`.
- `._` files ignored in `Invoke-SeriesEpisodeRename`.
- New `Get-SeriesEpisodeFromFileName` (S04E21, S01 Ep02, s1.e3, S02 Episode 4, 4x21). `Get-SeriesTitleClassification` keeps episode numbers from file names (Source `Name`), reserved up front, so a missing episode leaves a gap (Parks S4 has no E20) rather than shifting later numbers. A name number is ignored and flagged when it belongs to another season, is already used, or is claimed by two files. A named file is never a play-all.
- CHANGELOG and rename-series.ps1 help updated.

**Real F: result:** the series was ALREADY renamed at 18:47-18:54 on 2026-10-21 (by the user, from the fix branch build) and the S1/S2/S3/S6 folders became `Parks and Recreation Season 1/2/3/6`. The post-merge `-Apply -Yes` renamed 0 (all 112 already final). Every `rename-manifest.csv` row in all 6 folders was checked: NewName matches the S##/E## in OriginalName, no mismatches; Season 4 ends E19, E21, E22. Each folder has a manifest and `undo-rename.ps1`.

**Open on `F:\Series\Parks and Recreation` (need user decision; media drive NOT touched):**
- (a) Orphaned Jellyfin `*-thumb.jpg` under old names (S2 25, S3 17, S4 21, S5 22) no longer match the renamed videos.
- (b) `._` AppleDouble files remain in Season 4 (21) and season 5 (22).
- (c) Folder `Parks and Recreation season 5` is lowercase (cosmetic).
- (d) Season 6 long episodes flagged "a special?" but kept as episodes: Ep1 43:05, Ep16 1:04:47, Ep19 1:04:35, Ep20 53:03 (probably double/extended; unverified, TMDb key still unset).

### 2026-10-11 - watch-drives.ps1, insert-to-start (branch `feature/watch-drives-insert-to-start`)

`watch-drives.ps1` listens for `Win32_VolumeChangeEvent` (EventType 2) via
`Register-CimIndicationEvent` and opens `rip-disc.ps1 -Drive X:` in a new window for a
DVD-Video or Blu-ray. Idea borrowed from Automatic Ripping Machine (udev on Linux).
`rip-disc.ps1` itself is untouched and still interactive.

- **Antivirus (AVG is the active AV here; Defender is off):** do NOT start PowerShell with
  `-EncodedCommand`, least of all windowless, from scripts or tests. The first draft did;
  AVG's behaviour shield deleted `tests\Test-WatchDrives.ps1` mid-run and that exact path
  now gives "access denied" on create. The tests live in `tests\Test-DriveWatcher.ps1` and
  start no child PowerShell. The new window is launched as
  `-NoExit -File watch-drives.ps1 -RunRipFor X: -LockPath <file>`; extra rip arguments
  travel in the lock file, which `-RunRipFor` deletes when `rip-disc.ps1` returns
  (that deletion is how the watcher knows the drive is free again).
- **Tests:** `Test-DriveWatcher.ps1` 38/38. Listener registers unelevated (smoke-tested).
- **Not done:** no real disc inserted with it running; audio-CD event behaviour unknown;
  the real new-window launch has not been seen on screen.

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
