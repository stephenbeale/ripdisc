# RipDisc Project Roadmap

## Feature - -extras suffixed on title should just put all those files in the existing film's title
e.g. Fame-extras should look for an existing Fame dir and, if extras dir exists within, use that as output dir, else make that dir e.g. Fame/extras and place all final files there, only rename with title prefix BUT leave file names as they are after that.

## Feature - Check for dir char length
Handle all output max. char lengths so that they do not break- warn user of this and offer to abort to allow them to input a shorter title. Consider all sub dirs.

## Feature - Standalone series episode rename utility (added 2026-10-04) - FIRST VERSION BUILT 2026-10-05
Built as `rename-series.ps1` + `SeriesRetroRename.ps1` (dry run by default, `-Apply` to rename; see README "Additional Tools"). Still open: TheDiscDB use (needs a disc hash), episode titles in names, recording a manifest for folders hand-renamed earlier with no manifest, and a real-media test.
**Status 2026-10-05:** PR #151 is open (not merged) and `-Apply` is UNTESTED on real media. Next: run `-Apply` on the safe copy `C:\Video\Series\Joking Apart` (verify rename-manifest.csv, `undo-rename.ps1 -WhatIf`, then undo), then on `F:\Series\Joking Apart`. Concern: the dry run planned S02 Disc 2 `D1_t00` as S02-E07, probably wrong (likely a pilot or an extra) - investigate the extras/episode heuristics. Ideas: optional TMDb key for runtime-based extras detection in the standalone tool, and a `-Specials`/S00 option (see PR #152).
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

## Feature - Classify extras and move them out as part of the series rename (added 2026-10-04)
Part of the standalone utility above: while renaming a series folder, classify each file as episode or
extra using the criteria ALREADY in place, then move extras out using the existing move logic, so a
retroactive rename ends up laid out exactly as a fresh `-Series` rip. Since 2026-10-05 extras go to
the series-level `Specials` folder, not `DiscN\extras\`; the notes below describe the PR #143 design.

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

### Real-disc end-to-end test of series naming (#140/#142/#143) - PARTLY DONE 2026-10-05
The user confirmed `S##-E##` episode renaming works on a real disc (Silicon Valley). Still
unconfirmed against real hardware/APIs: the Disc 2+ start-episode prompt, extras detection, a live
TMDb lookup, and a real TheDiscDB hash match - ideally a Blu-ray likely to already be in TheDiscDB's
catalogue (ripped with `-Drive X:` so the hash can actually be computed; see the TheDiscDB section of
the README for why `-DriveIndex` alone skips it). The new `<Series>\Specials\` extras location
(below) also needs a real rip.

### Jellyfin and series extras - DONE 2026-10-05 (moved to `<Series>\Specials\`)
Jellyfin did NOT pick up PR #143's `DiscN\extras\` (checked on Silicon Valley S03). Jellyfin's docs
only recognise extras folders at series or season level. The user decided that series extras go in a
`Specials` folder at series level, beside the `Season N` folders (Jellyfin treats it as Season 00),
named `<Title>-S##-D#-Extra##` so discs sharing the folder never collide. Still to check: how Jellyfin
lists `-S##-D#-Extra##` files inside Specials (as specials with no TMDb match), and moving extras
already sitting in old `DiscN\extras\` folders (the user moved Silicon Valley S03's by hand).

### Deferred: port series naming to the C# `RipDisc -processQueue` / `-Queue` path
The C# queue processor still has none of the `S##-E##` renaming, extras detection, TheDiscDB
lookup, `extras` subfolder, manifest, or undo from #140/#142/#143 - it writes series files straight
into the Season folder with the older `Title-S##-...` prefix. The user has explicitly deferred this
port; documented in README's "Known limitation" note under TV series episode naming.

### CLAUDE.md has grown to ~170 KB of session notes
`CLAUDE.md` is loaded into every session and is now almost entirely chronological incident/session
history rather than conventions. Worth trimming down to conventions-only, with the history moved to
a separate file (e.g. `docs/session-history.md` or similar). Not done in this pass - logged only, at
the user's request, so it can be planned deliberately rather than rushed alongside a docs cleanup.

### Series flow: episode titles in names (idea, 2026-10-04)
Series renames currently produce `S##-E##` names only. Consider pulling episode titles (TheTVDB DVD
order, or TheDiscDB where it has them) into the name, e.g. `Show - S03E01 - Title.mp4`. Found while
hand-renaming Silicon Valley S03 Disc1. PAL rips run ~4% shorter than listed runtimes, so any
duration-based matching needs a 25fps tolerance.

### Series flow: retro-rename an already-ripped DiscN folder (idea, 2026-10-04)
Allow an existing `DiscN` folder (raw MakeMKV names) to be run through the Series naming flow after
the fact, writing the usual `rename-manifest.csv` and `undo-rename.ps1`. Motivation: Silicon Valley
S01 Disc2 and S03 Disc2 were ripped before the series flow existed, and one earlier rename left no
manifest, so the original disc names were lost.
