# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

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
