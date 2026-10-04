# RipDisc Project Roadmap

## Feature - -extras suffixed on title should just put all those files in the existing film's title
e.g. Fame-extras should look for an existing Fame dir and, if extras dir exists within, use that as output dir, else make that dir e.g. Fame/extras and place all final files there, only rename with title prefix BUT leave file names as they are after that.

## Feature - Check for dir char length
Handle all output max. char lengths so that they do not break- warn user of this and offer to abort to allow them to input a shorter title. Consider all sub dirs.

## Feature - tag bluray rips without affecting file naming for Jellyfin
Blu-ray args already passed but this append -BluRay onto file name after existing 'Feature' suffix, this would then allow me to identify BR version in jellfyin

## Feature - Standalone series episode rename utility (added 2026-10-04)
A retroactive, bulk tool for series that are ALREADY ripped on disk (e.g. `F:\Series\...`), not tied
to a live rip. The user has many series folders of episodes to rename, and it must be done safely
while keeping the original file names.

Requirements:
- Bulk: point it at a series root (or a Season/DiscN folder) and have it walk every folder, handling
  each DiscN folder as its own unit.
- Dry-run/preview by default; nothing is touched without an explicit apply switch (and per-folder
  confirmation, as in `Confirm-SeriesRenamePlan`).
- Never overwrite: a target that already exists is skipped with a warning (same rule as
  `Invoke-SeriesRenamePlan` and `undo-rename.ps1`). Already-renamed files are detected and left alone
  (`Get-SeriesNamePattern`).
- Keep original names: for EVERY folder renamed, write a ripdisc-format `rename-manifest.csv`
  (`Write-RenameManifest`: OriginalName, NewName, Kind, OriginalPath, NewPath, Timestamp) BEFORE any
  file is touched, and copy `undo-rename.ps1` next to it, so original disc/file names are recorded
  and the rename is fully reversible. This is a standing convention for all media renames.
- Reuse, do not reinvent: `SeriesEpisodes.ps1` (`Get-SeriesTitleClassification`,
  `New-SeriesRenamePlan`, `Show-SeriesRenamePlan`, `Write-RenameManifest`, `Invoke-SeriesRenamePlan`,
  `Invoke-SeriesEpisodeRename`) already does this for a single DiscN folder during a rip; the utility
  is mainly a driver that runs it over many existing folders, plus TMDb/TheDiscDB lookups where
  available.
- Open idea (merged here, from the Silicon Valley rename work): optionally pull real episode titles
  (TMDb) into the new names, and handle folders that were renamed earlier by hand/other tools with no
  manifest (record a manifest from current names so a future undo is possible).

## Feature - Classify extras and move them to an `extras` subfolder as part of the series rename (added 2026-10-04)
Part of the standalone utility above: while renaming a series folder, classify each file as episode or
extra using the criteria ALREADY in place, then move extras into the `extras` sub-dir using the
existing move logic, so a retroactive rename ends up laid out exactly as a fresh `-Series` rip.

Reuse (no new criteria):
- Classification: `Get-SeriesTitleClassification` in `SeriesEpisodes.ps1` (play-all/composite 70-130%
  of the sum of the others, TMDb runtime match / under 60% of expected runtime, median heuristic
  under 60% of median, TheDiscDB map, and the `-Overrides` edit option). Naming via
  `Get-SeriesExtraFileName` / `Get-SeriesExtraLabel`.
- Move: `New-SeriesRenamePlan` sets extras' NewName to `extras\<name>` (`$script:SeriesExtrasFolder`
  = `'extras'`), and `Invoke-SeriesRenamePlan` creates the folder on first use only and uses
  `Move-Item` without `-Force` (never overwrites).
- Reversibility: extras are recorded in the same `rename-manifest.csv` as `extras\<name>`, so
  `undo-rename.ps1` moves them back out and removes the empty `extras` folder.
- Files already inside an existing `extras` folder keep their numbers (as in `Invoke-SeriesEpisodeRename`).

## Backlog / Open Items (added 2026-10-04, after PRs #140/#142/#143)

### Real-disc end-to-end test of series naming (#140/#142/#143) - not yet done
Rip Disc 1 then Disc 2 of a real series season and check, against actual hardware/APIs rather
than fixtures: the Disc 2+ start-episode prompt, extras detection, a live TMDb lookup, and a real
TheDiscDB hash match - ideally a Blu-ray likely to already be in TheDiscDB's catalogue (ripped with
`-Drive X:` so the hash can actually be computed; see the TheDiscDB section of the README for why
`-DriveIndex` alone skips it). None of PRs #140/#142/#143 have been exercised against a real rip -
only AST-extracted logic tests and fixtures so far.

### Confirm Jellyfin recognises `DiscN\extras\` as extras under a non-season folder
PR #143 moved series extras into `DiscN\extras\` (per disc, same lowercase folder name movie mode
uses). Jellyfin's handling of an `extras\` folder nested under a plain `DiscN` folder (not a
`Season N` folder) has not been verified - confirm it's picked up as extras rather than ignored or
miscategorised.

### Deferred: port series naming to the C# `RipDisc -processQueue` / `-Queue` path
The C# queue processor still has none of the `S##-E##` renaming, extras detection, TheDiscDB
lookup, `extras` subfolder, manifest, or undo from #140/#142/#143 - it writes series files straight
into the Season folder with the older `Title-S##-...` prefix. The user has explicitly deferred this
port; documented in README's "Known limitation" note under TV series episode naming.

### PR #141 (MakeMKV progress bar) likely needs a rebase before merging
Open PR #141 (`feature/makemkv-progress`, from another session) adds a MakeMKV progress bar, 10%
milestones and ETA to `rip-disc.ps1`. It was branched before #140/#142/#143 landed and almost
certainly conflicts with them in `rip-disc.ps1` (all four PRs touch the same rip/organize code).
Needs a rebase onto `main` before it can merge cleanly. Not touched as part of this pass - logged
here only.

### CLAUDE.md has grown to ~170 KB of session notes
`CLAUDE.md` is loaded into every session and is now almost entirely chronological incident/session
history rather than conventions. Worth trimming down to conventions-only, with the history moved to
a separate file (e.g. `docs/session-history.md` or similar). Not done in this pass - logged only, at
the user's request, so it can be planned deliberately rather than rushed alongside a docs cleanup.
