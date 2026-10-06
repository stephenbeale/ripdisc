<#
.SYNOPSIS
    Renames series episodes that are ALREADY ripped on disk to <Title>-S##-E##, moving
    extras into the series-level Specials folder, with a manifest and undo script for every folder.

.DESCRIPTION
    The retroactive twin of a live `rip-disc.ps1 -Series` rip: same naming, same episode vs
    extra detection (TMDb runtimes, median length, play-all detection), same confirmation
    table, same rename-manifest.csv + undo-rename.ps1 - it simply runs them over folders
    that already exist (SeriesEpisodes.ps1 does the work; this is the driver).

    SAFE BY DEFAULT: without -Apply it only shows the plan for every folder and changes
    nothing (-WhatIf means the same dry run, even alongside -Apply). With -Apply, each
    folder shows its table and asks (Enter accepts, n leaves the folder alone, e edits
    episode/extra); -Yes accepts every table without asking.
    Before any file in a folder moves, rename-manifest.csv (OriginalName, NewName, Kind,
    OriginalPath, NewPath, Timestamp) is written there and undo-rename.ps1 is copied next
    to it, so the original file names are always recoverable. Nothing is ever overwritten,
    and files already named <Title>-S##-E##/-Extra## are left alone.

    Episode numbers continue across the discs of a season (Disc2 starts after Disc1's last
    episode), whether the earlier discs are already renamed or only planned in this run.

    Not touched: the optical drive (this only reads file names and video durations).
    TheDiscDB needs the disc's hash from the physical disc, so it is not used here;
    classification uses TMDb (if a key is configured) or title lengths.

.PARAMETER Path
    A series folder (<root>\Silicon Valley), one Season folder, or one DiscN folder
    ("Disc 2" and "<Title>-Disc 2" folders are recognised too).

.PARAMETER Title
    Series title used in the new names. Default: the series folder's name.

.PARAMETER Season
    Season number when the path has no "Season N" folder (default: S01 fallback).

.PARAMETER StartEpisode
    First episode number for the first folder of each season (default 1). Later discs
    continue automatically.

.PARAMETER MarkSpecial
    File name wildcard(s), matched against the ORIGINAL names, to make specials
    (<Title>-S00-E##, in the series Specials folder), whatever their length says.
    Titles 1.6x longer than expected are flagged as possible specials, but only this
    (or the prompt's edit, 2s) makes one.

.PARAMETER MarkExtra
    File name wildcard(s) to make extras, whatever their length says.

.PARAMETER MarkEpisode
    File name wildcard(s) to make episodes, whatever their length says.

.PARAMETER Apply
    Actually rename. Without it, dry run only.

.PARAMETER Yes
    With -Apply: accept each folder's table without asking.

.PARAMETER NoTmdb
    Do not look up TMDb runtimes (use title lengths only).

.EXAMPLE
    .\rename-series.ps1 "F:\Series\Silicon Valley"                 # preview everything

.EXAMPLE
    .\rename-series.ps1 "F:\Series\Silicon Valley\Season 3" -Apply # rename Season 3, confirming each disc

.EXAMPLE
    .\rename-series.ps1 "F:\Series\Silicon Valley\Season 3\Disc2" -StartEpisode 6 -Apply

.EXAMPLE
    .\rename-series.ps1 "F:\Series\Boys from the Black Stuff\Season 1" -MarkSpecial "*Disc 1 - E01*" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$Path,
    [string]$Title = "",
    [int]$Season = 0,
    [int]$StartEpisode = 0,
    [switch]$Apply,
    [switch]$Yes,
    [switch]$NoTmdb,
    [string[]]$MarkSpecial = @(),
    [string[]]$MarkExtra = @(),
    [string[]]$MarkEpisode = @()
)

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Error "Folder not found: $Path"
    exit 1
}
if ($WhatIfPreference) {
    # -WhatIf is the dry run this script already does without -Apply.
    if ($Apply) { Write-Host "-WhatIf given: dry run only, -Apply ignored." -ForegroundColor Yellow }
    $Apply = $false
}
if ($Yes -and -not $Apply) {
    Write-Host "Note: -Yes only matters with -Apply; this is a dry run." -ForegroundColor DarkGray
}

. (Join-Path $PSScriptRoot "Load-Config.ps1")
. (Join-Path $PSScriptRoot "SeriesEpisodes.ps1")
. (Join-Path $PSScriptRoot "SeriesRetroRename.ps1")

$script:LogFile = Join-Path $env:TEMP ("ripdisc-rename-series-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
function Write-Log {
    param([string]$Message)
    if ($Apply) { Add-Content -LiteralPath $script:LogFile -Value ("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) }
}

$markKinds = @{}
foreach ($m in $MarkEpisode) { $markKinds[$m] = 'Episode' }
foreach ($m in $MarkExtra) { $markKinds[$m] = 'Extra' }
foreach ($m in $MarkSpecial) { $markKinds[$m] = 'Special' }

$summary = Invoke-SeriesRetroRename -Root $Path -Title $Title -Season $Season -StartEpisode $StartEpisode `
    -Apply:$Apply -Yes:$Yes -NoTmdb:$NoTmdb -MarkKinds $markKinds `
    -HandBrakePath $script:Config_HandBrakePath `
    -UndoScriptSource (Join-Path $PSScriptRoot "undo-rename.ps1")
if ($Apply -and $summary.Units -gt 0) { Write-Host "Log: $script:LogFile" -ForegroundColor Gray }
