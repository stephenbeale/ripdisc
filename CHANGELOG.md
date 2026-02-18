# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

## 2026-02-18

### Added
- `-Tutorial` and `-Fitness` disc type switches — output to `E:\Tutorials\` and `E:\Fitness\` (#40)
- Extras disc: encode directly into `extras\` subdirectory instead of encoding then moving (#43)
- Extras disc: files prefixed with title only, no `-extras` or `-Special Features` in name (#42)
- Empty parent directory cleanup after temp directory removal (#39)

### Fixed
- Series concurrent disc rename conflicts — use per-disc subdirs for encoding isolation (#41)

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
