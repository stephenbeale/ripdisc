#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for rename-series.ps1's engine (SeriesRetroRename.ps1): folder discovery, dry run
    that changes nothing, apply with manifest + undo, numbering across discs, idempotence.

.DESCRIPTION
    Dot-sources the real SeriesEpisodes.ps1 and SeriesRetroRename.ps1 - nothing is
    reimplemented. Video durations are injected (-GetDuration) and TMDb is replaced by an
    injected lookup, so no network, no video tools, no disc and no optical drive. Filesystem
    cases use a temp directory of empty placeholder files, removed afterwards. Nothing under
    a real media drive is touched.

.EXAMPLE
    .\tests\Test-SeriesRetroRename.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$undoPath = Join-Path $repoRoot 'undo-rename.ps1'
. (Join-Path $repoRoot 'SeriesEpisodes.ps1')
. (Join-Path $repoRoot 'SeriesRetroRename.ps1')

function Write-Log { param([string]$Message) }

$script:Passed = 0
$script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}
function Assert-True {
    param([bool]$Condition, [string]$Because)
    if ($Condition) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because" -ForegroundColor Red }
}

# Video files below a folder as relative paths (extras\ included), sorted.
function Get-RelNames {
    param([string]$Dir)
    $root = (Resolve-Path -LiteralPath $Dir).ProviderPath.TrimEnd('\') + '\'
    (Get-ChildItem -LiteralPath $Dir -File -Recurse | Where-Object { $_.Extension -in '.mp4', '.mkv' } |
        ForEach-Object { $_.FullName.Substring($root.Length) } | Sort-Object) -join ','
}

# Builds a disc folder of empty files and records each one's pretend duration (minutes).
$script:Durations = @{}
function New-DiscFolder {
    param([string]$Dir, [hashtable]$FilesAndMinutes)
    New-Item -ItemType Directory -Path $Dir -Force | Out-Null
    foreach ($name in $FilesAndMinutes.Keys) {
        New-Item -ItemType File -Path (Join-Path $Dir $name) | Out-Null
        $script:Durations[(Join-Path $Dir $name)] = $FilesAndMinutes[$name] * 60
    }
}
$getDuration = { param($path) $script:Durations[$path] }

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-retro-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nFolder discovery" -ForegroundColor Cyan

    $series = Join-Path $tempRoot 'Show A'
    New-DiscFolder (Join-Path $series 'Season 2\Disc2') @{ 'a_t00.mkv' = 25 }
    New-DiscFolder (Join-Path $series 'Season 2\Disc10') @{ 'a_t00.mkv' = 25 }
    New-DiscFolder (Join-Path $series 'Season 2\Disc1') @{ 'a_t00.mkv' = 25 }
    New-DiscFolder (Join-Path $series 'Season 1\Disc1') @{ 'a_t00.mkv' = 25 }
    New-Item -ItemType Directory -Path (Join-Path $series 'Season 2\Disc1\extras') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $series 'Season 2\Notes') | Out-Null

    $units = @(Get-SeriesRenameUnits -Root $series)
    Assert-Equal 'S1D1,S2D1,S2D2,S2D10' (($units | ForEach-Object { "S$($_.Season)D$($_.Disc)" }) -join ',') 'series root: seasons in order, discs in NUMERIC order (Disc10 after Disc2), extras/other folders ignored'
    Assert-True (@($units | Where-Object { $_.Title -ne 'Show A' }).Count -eq 0) 'title comes from the series folder name'

    $units = @(Get-SeriesRenameUnits -Root (Join-Path $series 'Season 2'))
    Assert-Equal 'Show A' $units[0].Title 'season folder: title is the folder above'
    Assert-Equal '2' "$($units[0].Season)" 'season folder: season parsed from "Season 2"'
    Assert-Equal 3 $units.Count 'season folder: its three disc folders'

    $units = @(Get-SeriesRenameUnits -Root (Join-Path $series 'Season 2\Disc2'))
    Assert-Equal 1 $units.Count 'disc folder: exactly one unit'
    Assert-Equal 'Show A|2|2' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") 'disc folder: title/season/disc inferred from the path'

    $units = @(Get-SeriesRenameUnits -Root $series -Title 'Better Name')
    Assert-True (@($units | Where-Object { $_.Title -ne 'Better Name' }).Count -eq 0) '-Title overrides the folder name'

    $noSeason = Join-Path $tempRoot 'Fargo'
    New-DiscFolder (Join-Path $noSeason 'Disc1') @{ 'f_t00.mkv' = 40 }
    $units = @(Get-SeriesRenameUnits -Root $noSeason)
    Assert-Equal 'Fargo|0|1' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") 'no Season folder: Disc folder directly under the series, season 0 (S01 fallback)'

    $flat = Join-Path $tempRoot 'Flat\Season 3'
    New-DiscFolder $flat @{ 'x.mkv' = 30 }
    $units = @(Get-SeriesRenameUnits -Root $flat)
    Assert-Equal 'Flat|3|0' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") 'Season folder with files and no Disc folders is itself the unit'

    # ---------------------------------------------------------------------------
    Write-Host "`nDry run changes nothing and numbers across discs" -ForegroundColor Cyan

    $show = Join-Path $tempRoot 'Silicon Test'
    $d1 = Join-Path $show 'Season 3\Disc1'
    $d2 = Join-Path $show 'Season 3\Disc2'
    New-DiscFolder $d1 @{ 'SV_t00.mkv' = 27; 'SV_t01.mkv' = 28; 'SV_t02.mkv' = 27; 'SV_t03.mkv' = 5 }
    New-DiscFolder $d2 @{ 'SV_t00.mkv' = 28; 'SV_t01.mkv' = 27 }
    $before1 = Get-RelNames $d1
    $before2 = Get-RelNames $d2

    $s = Invoke-SeriesRetroRename -Root $show -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal $before1 (Get-RelNames $d1) 'dry run: Disc1 file names unchanged'
    Assert-Equal $before2 (Get-RelNames $d2) 'dry run: Disc2 file names unchanged'
    Assert-True (-not (Test-Path (Join-Path $d1 'rename-manifest.csv'))) 'dry run: no manifest written'
    Assert-True (-not (Test-Path (Join-Path $d1 'undo-rename.ps1'))) 'dry run: no undo script copied'
    Assert-True (-not (Test-Path (Join-Path $d1 'extras'))) 'dry run: no extras folder created'
    Assert-Equal 6 $s.Planned 'dry run: reports 6 files planned'
    Assert-Equal 0 $s.Renamed 'dry run: reports 0 renamed'
    $plan2 = @($s.Results[1].Result.Plan | Where-Object { $_.Kind -eq 'Episode' } | ForEach-Object { $_.NewName }) -join ','
    Assert-Equal 'Silicon Test-S03-E04.mkv,Silicon Test-S03-E05.mkv' $plan2 'dry run: Disc2 planned from E04 (continues after the PLANNED Disc1 episodes E01-E03)'
    Assert-Equal 'extras\Silicon Test-S03-Extra01.mkv' (@($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Extra' })[0].NewName) 'dry run: the 5-minute title is planned as extras\...-Extra01'

    # ---------------------------------------------------------------------------
    Write-Host "`nApply: rename, manifest before move, undo copied" -ForegroundColor Cyan

    $s = Invoke-SeriesRetroRename -Root $show -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'extras\Silicon Test-S03-Extra01.mkv,Silicon Test-S03-E01.mkv,Silicon Test-S03-E02.mkv,Silicon Test-S03-E03.mkv' (Get-RelNames $d1) 'apply: Disc1 episodes renamed, extra moved to extras\'
    Assert-Equal 'Silicon Test-S03-E04.mkv,Silicon Test-S03-E05.mkv' (Get-RelNames $d2) 'apply: Disc2 continues at E04'
    Assert-Equal 6 $s.Renamed 'apply: reports 6 renamed'
    foreach ($d in @($d1, $d2)) {
        $leaf = Split-Path $d -Leaf
        Assert-True (Test-Path (Join-Path $d 'rename-manifest.csv')) "apply: $leaf has rename-manifest.csv"
        Assert-True (Test-Path (Join-Path $d 'undo-rename.ps1')) "apply: $leaf has undo-rename.ps1 copied in"
    }
    $rows = @(Import-Csv -LiteralPath (Join-Path $d1 'rename-manifest.csv'))
    Assert-Equal 'OriginalName,NewName,Kind,OriginalPath,NewPath,Timestamp' ($rows[0].PSObject.Properties.Name -join ',') 'manifest uses the ripdisc column format'
    Assert-Equal 4 $rows.Count 'Disc1 manifest has a row per renamed file'
    Assert-Equal 'SV_t00.mkv' $rows[0].OriginalName 'manifest records the ORIGINAL file name'
    Assert-Equal 'extras\Silicon Test-S03-Extra01.mkv' (($rows | Where-Object { $_.Kind -eq 'Extra' }).NewName) 'manifest records the extra as extras\<name>'

    # ---------------------------------------------------------------------------
    Write-Host "`nRe-running is a no-op (already-renamed files left alone)" -ForegroundColor Cyan

    $afterApply1 = Get-RelNames $d1
    $manifestLines = (Get-Content (Join-Path $d1 'rename-manifest.csv')).Count
    $s = Invoke-SeriesRetroRename -Root $show -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 0 $s.Renamed 're-run: nothing renamed'
    Assert-Equal $afterApply1 (Get-RelNames $d1) 're-run: names unchanged'
    Assert-Equal $manifestLines (Get-Content (Join-Path $d1 'rename-manifest.csv')).Count 're-run: manifest not appended to'

    # ---------------------------------------------------------------------------
    Write-Host "`nUndo restores the original disc names" -ForegroundColor Cyan

    & (Join-Path $d1 'undo-rename.ps1') *> $null
    & (Join-Path $d2 'undo-rename.ps1') *> $null
    Assert-Equal $before1 (Get-RelNames $d1) 'undo: Disc1 back to original names (extras folder emptied)'
    Assert-Equal $before2 (Get-RelNames $d2) 'undo: Disc2 back to original names'
    Assert-True (-not (Test-Path (Join-Path $d1 'extras'))) 'undo: empty extras folder removed'

    # ---------------------------------------------------------------------------
    Write-Host "`nDeclining leaves a folder alone; a single Disc2 can continue from renamed Disc1" -ForegroundColor Cyan

    Remove-Item (Join-Path $d1 'rename-manifest.csv'), (Join-Path $d2 'rename-manifest.csv') -Force
    $answers = New-Object System.Collections.Queue
    'n' | ForEach-Object { $answers.Enqueue($_) }
    $readNo = { param($p) if ($answers.Count -gt 0) { $answers.Dequeue() } else { $null } }.GetNewClosure()
    $s = Invoke-SeriesRetroRename -Root $d1 -Apply -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath -ReadInput $readNo 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 1 $s.Declined 'answering n declines the folder'
    Assert-Equal $before1 (Get-RelNames $d1) 'declined: files keep their names'
    Assert-True (-not (Test-Path (Join-Path $d1 'rename-manifest.csv'))) 'declined: no manifest written'

    # rename Disc1 for real, then point at Disc2 alone: it should start after Disc1's last episode.
    $null = Invoke-SeriesRetroRename -Root $d1 -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1
    $s = Invoke-SeriesRetroRename -Root $d2 -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Silicon Test-S03-E04.mkv,Silicon Test-S03-E05.mkv' (@($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ',') 'Disc2 alone starts at E04 from the renamed Disc1 on disk'

    # ---------------------------------------------------------------------------
    Write-Host "`n-StartEpisode and no earlier discs" -ForegroundColor Cyan

    $lone = Join-Path $tempRoot 'Lone\Season 1\Disc2'
    New-DiscFolder $lone @{ 'L_t00.mkv' = 40; 'L_t01.mkv' = 41 }
    $s = Invoke-SeriesRetroRename -Root $lone -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Lone-S01-E01.mkv,Lone-S01-E02.mkv' (@($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ',') 'Disc2 with no earlier discs and no -StartEpisode starts at E01 (with a warning)'
    $s = Invoke-SeriesRetroRename -Root $lone -NoTmdb -StartEpisode 7 -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Lone-S01-E07.mkv,Lone-S01-E08.mkv' (@($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ',') '-StartEpisode 7 numbers from E07'

    # ---------------------------------------------------------------------------
    Write-Host "`nNever overwrites" -ForegroundColor Cyan

    $clash = Join-Path $tempRoot 'Clash\Season 1\Disc1'
    New-DiscFolder $clash @{ 'C_t00.mkv' = 40; 'C_t01.mkv' = 41 }
    $target = Join-Path $clash 'Clash-S01-E01.mkv'
    $s = Invoke-SeriesRetroRename -Root $clash -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Clash-S01-E01.mkv,Clash-S01-E02.mkv' (@($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ',') 'baseline plan for the clash folder'
    Set-Content -LiteralPath $target -Value 'precious' -Encoding ASCII
    # E01 now exists, so it counts as already renamed (taken): the two originals must take
    # E02 and E03, and the existing file is neither touched nor reused.
    $null = Invoke-SeriesRetroRename -Root $clash -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1
    Assert-Equal 'precious' ((Get-Content -LiteralPath $target -Raw).Trim()) 'an existing target file is left byte-for-byte untouched'
    Assert-Equal 'Clash-S01-E01.mkv,Clash-S01-E02.mkv,Clash-S01-E03.mkv' (Get-RelNames $clash) 'originals took the next free numbers instead of overwriting'

    # ---------------------------------------------------------------------------
    Write-Host "`nOlder rip layouts (Joking Apart): Series N folders, several discs in one folder" -ForegroundColor Cyan

    Assert-Equal '1' "$(Get-SeasonFolderNumber 'Series 1')" 'season name: "Series 1"'
    Assert-Equal '2' "$(Get-SeasonFolderNumber 'Season 2')" 'season name: "Season 2"'
    Assert-Equal '1' "$(Get-SeasonFolderNumber 'Joking Apart-Series 1')" 'season name: "<anything>-Series 1"'
    Assert-Equal '3' "$(Get-SeasonFolderNumber 'Joking Apart Series 3')" 'season name: "<anything> Series 3"'
    Assert-Equal '4' "$(Get-SeasonFolderNumber 'Joking Apart-Season 4')" 'season name: "<anything>-Season 4"'
    Assert-Equal '12' "$(Get-SeasonFolderNumber 'show - SERIES 12')" 'season name: case-insensitive, two digits'
    Assert-Equal '5' "$(Get-SeasonFolderNumber 'series5')" 'season name: no space'
    Assert-True ($null -eq (Get-SeasonFolderNumber 'Joking Apart')) 'a plain series folder is not a season'
    Assert-True ($null -eq (Get-SeasonFolderNumber 'Series 1 Extras')) 'a folder merely containing "Series 1" is not a season'
    Assert-Equal '2' "$(Get-DiscFolderNumber 'Disc 2')" 'disc folder: "Disc 2" (with space)'
    Assert-Equal '3' "$(Get-DiscFolderNumber 'disc3')" 'disc folder: "disc3"'
    Assert-Equal '1' "$(Get-DiscNumberFromFileName 'Joking Apart-Series 2 Disc 1-B1_t06.mp4')" 'file token: "Series 2 Disc 1-..."'
    Assert-True ($null -eq (Get-DiscNumberFromFileName 'Joking Apart-Series 1-C1_t00.mp4')) 'file token: none when the name has no Disc N'

    $ja = Join-Path $tempRoot 'Joking Apart'
    $ja1 = Join-Path $ja 'Joking Apart-Series 1'
    $ja2 = Join-Path $ja 'Joking Apart-Series 2'
    $s1Files = @{}
    foreach ($n in 'C1_t00','C1_t01','C1_t02','C1_t03','C1_t04','C1_t06') { $s1Files["Joking Apart-Series 1-$n.mp4"] = 30 }
    New-DiscFolder $ja1 $s1Files
    New-DiscFolder (Join-Path $ja1 'extras') @{ 'Joking Apart-Series 1-B1_t05.mp4' = 4 }
    $s2Files = @{ 'Joking Apart-Series 2 Disc 1-B1_t06.mp4' = 5 }
    foreach ($n in 1..6) { $s2Files["Joking Apart-Series 2 Disc 1-C$($n)_t0$($n - 1).mp4"] = 30 }
    $s2Files['Joking Apart-Series 2 Disc 2-D1_t00.mp4'] = 30
    New-DiscFolder $ja2 $s2Files
    $beforeJa1 = Get-RelNames $ja1
    $beforeJa2 = Get-RelNames $ja2

    $units = @(Get-SeriesRenameUnits -Root $ja)
    Assert-Equal 'S1D0,S2D1,S2D2' (($units | ForEach-Object { "S$($_.Season)D$($_.Disc)" }) -join ',') 'series root: Series 1 is one unit, Series 2 splits into Disc 1 and Disc 2'
    Assert-True (@($units | Where-Object { $_.Title -ne 'Joking Apart' }).Count -eq 0) 'title is the series folder (Joking Apart), not the season folder name'
    Assert-True ($null -eq $units[0].FileFilter -and $null -ne $units[1].FileFilter) 'no disc tokens -> no filter; disc tokens -> a filter per disc'
    Assert-True ($units[1].Directory -eq $ja2 -and $units[2].Directory -eq $ja2) 'both Series 2 units stay in the one folder (nothing moved into Disc folders)'
    Assert-True (@($units | Where-Object { $_.Directory -like '*extras*' }).Count -eq 0) 'an existing extras folder is never a unit'

    $units = @(Get-SeriesRenameUnits -Root $ja1)
    Assert-Equal 'Joking Apart|1|0' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") 'pointing straight at "Joking Apart-Series 1": title is the parent folder'
    $units = @(Get-SeriesRenameUnits -Root $ja2)
    Assert-Equal 2 $units.Count 'pointing straight at Series 2: two disc units'

    $spaced = Join-Path $tempRoot 'Spaced\Series 3\Disc 2'
    New-DiscFolder $spaced @{ 's_t00.mkv' = 40 }
    $units = @(Get-SeriesRenameUnits -Root (Join-Path $tempRoot 'Spaced'))
    Assert-Equal 'Spaced|3|2' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") '"Series 3\Disc 2" (space) is found from the series root'
    $units = @(Get-SeriesRenameUnits -Root $spaced)
    Assert-Equal 'Spaced|3|2' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") '"Disc 2" folder pointed at directly: season/title inferred'

    # Dry run
    $s = Invoke-SeriesRetroRename -Root $ja -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal $beforeJa1 (Get-RelNames $ja1) 'dry run: Series 1 untouched'
    Assert-Equal $beforeJa2 (Get-RelNames $ja2) 'dry run: Series 2 untouched'
    Assert-True (-not (Test-Path (Join-Path $ja2 'rename-manifest.csv'))) 'dry run: no manifest'
    $p1 = @($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ','
    Assert-Equal 'Joking Apart-S01-E01.mp4,Joking Apart-S01-E02.mp4,Joking Apart-S01-E03.mp4,Joking Apart-S01-E04.mp4,Joking Apart-S01-E05.mp4,Joking Apart-S01-E06.mp4' $p1 'Series 1: six files -> E01-E06 (existing extras\ file ignored)'
    $d1plan = @($s.Results[1].Result.Plan)
    $d2plan = @($s.Results[2].Result.Plan)
    Assert-Equal 7 $d1plan.Count 'Series 2 Disc 1 unit takes only the 7 "Disc 1" files'
    Assert-Equal 1 $d2plan.Count 'Series 2 Disc 2 unit takes only the "Disc 2" file'
    Assert-Equal 'extras\Joking Apart-S02-Extra01.mp4' (($d1plan | Where-Object { $_.Kind -eq 'Extra' }).NewName) 'the 5-minute Disc 1 title is planned as an extra'
    Assert-Equal 'Joking Apart-S02-E07.mp4' $d2plan[0].NewName 'Disc 2 continues numbering after Disc 1 (E07)'

    # Apply
    $s = Invoke-SeriesRetroRename -Root $ja -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'extras\Joking Apart-S02-Extra01.mp4,Joking Apart-S02-E01.mp4,Joking Apart-S02-E02.mp4,Joking Apart-S02-E03.mp4,Joking Apart-S02-E04.mp4,Joking Apart-S02-E05.mp4,Joking Apart-S02-E06.mp4,Joking Apart-S02-E07.mp4' (Get-RelNames $ja2) 'apply: Series 2 episodes E01-E07 across both discs, extra moved to extras\'
    Assert-True (Test-Path -LiteralPath (Join-Path $ja1 'extras\Joking Apart-Series 1-B1_t05.mp4')) 'apply: the pre-existing extras file is left exactly where it was, under its old name'
    $rows = @(Import-Csv -LiteralPath (Join-Path $ja2 'rename-manifest.csv'))
    Assert-Equal 8 $rows.Count 'two units in one folder: ONE manifest holding both units rows (appended, not clobbered)'
    Assert-Equal 8 @($rows | Select-Object -ExpandProperty OriginalName -Unique).Count 'manifest rows are distinct (every original name recorded once)'
    Assert-True (Test-Path (Join-Path $ja2 'undo-rename.ps1')) 'undo script copied next to the shared manifest'
    & (Join-Path $ja2 'undo-rename.ps1') *> $null
    Assert-Equal $beforeJa2 (Get-RelNames $ja2) 'undo of the shared manifest restores BOTH discs original names'
    & (Join-Path $ja1 'undo-rename.ps1') *> $null
    Assert-Equal $beforeJa1 (Get-RelNames $ja1) 'undo in Series 1 restores its names and keeps the existing extras file'

    # Two units sharing a folder never plan the same Extra number, and a later run continues
    $x = Join-Path $tempRoot 'Xtra\Xtra-Series 1'
    New-DiscFolder $x @{ 'Xtra-Series 1 Disc 1-A_t00.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t01.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t03.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t02.mkv' = 4; 'Xtra-Series 1 Disc 2-B_t00.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t01.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t03.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t02.mkv' = 4 }
    $s = Invoke-SeriesRetroRename -Root $x -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $xe1 = @($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Extra' } | ForEach-Object { $_.NewName }) -join ','
    $xe2 = @($s.Results[1].Result.Plan | Where-Object { $_.Kind -eq 'Extra' } | ForEach-Object { $_.NewName }) -join ','
    Assert-Equal 'extras\Xtra-S01-Extra01.mkv|extras\Xtra-S01-Extra02.mkv' "$xe1|$xe2" 'dry run: Disc 2 plans Extra02, not a second Extra01'

    $r = Join-Path $tempRoot 'Resume\Resume-Series 1'
    New-DiscFolder $r @{ 'Resume-S01-E01.mkv' = 30; 'Resume-S01-E02.mkv' = 30; 'Resume-Series 1 Disc 2-B_t00.mkv' = 30 }
    $s = Invoke-SeriesRetroRename -Root $r -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Resume-S01-E03.mkv' (@($s.Results[0].Result.Plan)[0].NewName) 'a Disc 2 file whose Disc 1 is already renamed in the same folder starts at E03'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "Passed: $script:Passed  Failed: $script:Failed" -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
if ($script:Failed -gt 0) { exit 1 }
