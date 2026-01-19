# RipDisc Project

PowerShell scripts for automated DVD and Blu-ray disc ripping using MakeMKV and HandBrake.

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
