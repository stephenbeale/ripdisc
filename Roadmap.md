# RipDisc Project Roadmap

## Feature - -extras suffixed on title should just put all those files in the existing film's title
e.g. Fame-extras should look for an existing Fame dir and, if extras dir exists within, use that as output dir, else make that dir e.g. Fame/extras and place all final files there, only rename with title prefix BUT leave file names as they are after that.

## Feature - Check for dir char length
Handle all output max. char lengths so that they do not break- warn user of this and offer to abort to allow them to input a shorter title. Consider all sub dirs.

## Feature - tag bluray rips without affecting file naming for Jellyfin
Blu-ray args already passed but this append -BluRay onto file name after existing 'Feature' suffix, this would then allow me to identify BR version in jellfyin

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
