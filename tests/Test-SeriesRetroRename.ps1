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

# Video files below a folder as relative paths (Specials\ / extras\ included), sorted.
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
    New-Item -ItemType Directory -Path (Join-Path $series 'Specials') | Out-Null

    $units = @(Get-SeriesRenameUnits -Root $series)
    Assert-Equal 'S1D1,S2D1,S2D2,S2D10' (($units | ForEach-Object { "S$($_.Season)D$($_.Disc)" }) -join ',') 'series root: seasons in order, discs in NUMERIC order (Disc10 after Disc2), Specials/extras/other folders ignored'
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
    Assert-True (-not (Test-Path (Join-Path $show 'Specials'))) 'dry run: no Specials folder created'
    Assert-Equal 6 $s.Planned 'dry run: reports 6 files planned'
    Assert-Equal 0 $s.Renamed 'dry run: reports 0 renamed'
    $plan2 = @($s.Results[1].Result.Plan | Where-Object { $_.Kind -eq 'Episode' } | ForEach-Object { $_.NewName }) -join ','
    Assert-Equal 'Silicon Test-S03-E04.mkv,Silicon Test-S03-E05.mkv' $plan2 'dry run: Disc2 planned from E04 (continues after the PLANNED Disc1 episodes E01-E03)'
    Assert-Equal '..\..\Specials\Silicon Test-S03-D1-Extra01.mkv' (@($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Extra' })[0].NewName) 'dry run: the 5-minute title is planned as <Series>\Specials\...-S03-D1-Extra01'

    # ---------------------------------------------------------------------------
    Write-Host "`nApply: rename, manifest before move, undo copied" -ForegroundColor Cyan

    $s = Invoke-SeriesRetroRename -Root $show -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Silicon Test-S03-E01.mkv,Silicon Test-S03-E02.mkv,Silicon Test-S03-E03.mkv' (Get-RelNames $d1) 'apply: Disc1 episodes renamed in place'
    Assert-Equal 'Silicon Test-S03-D1-Extra01.mkv' (Get-RelNames (Join-Path $show 'Specials')) 'apply: the extra moved to the series-level Specials folder'
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
    Assert-Equal '..\..\Specials\Silicon Test-S03-D1-Extra01.mkv' (($rows | Where-Object { $_.Kind -eq 'Extra' }).NewName) 'manifest records the extra as ..\..\Specials\<name>'

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
    Assert-Equal $before1 (Get-RelNames $d1) 'undo: Disc1 back to original names (extra moved back from Specials)'
    Assert-Equal $before2 (Get-RelNames $d2) 'undo: Disc2 back to original names'
    Assert-True (-not (Test-Path (Join-Path $show 'Specials'))) 'undo: emptied Specials folder removed'
    Assert-True (-not (Test-Path (Join-Path $d1 'rename-manifest.csv'))) 'undo: a fully undone manifest is retired, so the next rename starts a fresh one'
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $d1 -Filter 'rename-manifest.undone-*.csv').Count 'undo: ...and kept as rename-manifest.undone-<timestamp>.csv'

    # Re-apply after undo: the new manifest holds only the new rows (no stale ones to replay).
    $null = Invoke-SeriesRetroRename -Root $d1 -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1
    Assert-Equal 4 @(Import-Csv -LiteralPath (Join-Path $d1 'rename-manifest.csv')).Count 're-apply after undo: fresh manifest with only this run''s 4 rows'
    & (Join-Path $d1 'undo-rename.ps1') *> $null
    Assert-Equal $before1 (Get-RelNames $d1) 're-apply then undo: original names again'

    # ---------------------------------------------------------------------------
    Write-Host "`nDeclining leaves a folder alone; a single Disc2 can continue from renamed Disc1" -ForegroundColor Cyan

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
    Assert-Equal 'S1D0,S2D0' (($units | ForEach-Object { "S$($_.Season)D$($_.Disc)" }) -join ',') 'series root: each season folder is ONE unit, even when its files carry Disc 1 / Disc 2 tokens'
    Assert-True (@($units | Where-Object { $_.Title -ne 'Joking Apart' }).Count -eq 0) 'title is the series folder (Joking Apart), not the season folder name'
    Assert-True ($units[1].Directory -eq $ja2) 'the Series 2 unit is the folder itself (nothing moved into Disc folders)'
    Assert-True (@($units | Where-Object { $_.Directory -like '*extras*' }).Count -eq 0) 'an existing extras folder is never a unit'

    $units = @(Get-SeriesRenameUnits -Root $ja1)
    Assert-Equal 'Joking Apart|1|0' ("$($units[0].Title)|$($units[0].Season)|$($units[0].Disc)") 'pointing straight at "Joking Apart-Series 1": title is the parent folder'
    $units = @(Get-SeriesRenameUnits -Root $ja2)
    Assert-Equal 1 $units.Count 'pointing straight at Series 2: one unit for both discs'

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
    $s2plan = @($s.Results[1].Result.Plan)
    Assert-Equal 8 $s2plan.Count 'Series 2: one plan covering both discs (8 files)'
    Assert-Equal '..\Specials\Joking Apart-S02-Extra01.mp4' (($s2plan | Where-Object { $_.Kind -eq 'Extra' }).NewName) 'the 5-minute Disc 1 title is planned as an extra'
    Assert-Equal 'Joking Apart-S02-E07.mp4' (($s2plan | Where-Object { $_.OriginalName -like '*Disc 2*' }).NewName) 'the Disc 2 file comes after the Disc 1 episodes (E07)'

    # Apply
    $s = Invoke-SeriesRetroRename -Root $ja -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Joking Apart-S02-E01.mp4,Joking Apart-S02-E02.mp4,Joking Apart-S02-E03.mp4,Joking Apart-S02-E04.mp4,Joking Apart-S02-E05.mp4,Joking Apart-S02-E06.mp4,Joking Apart-S02-E07.mp4' (Get-RelNames $ja2) 'apply: Series 2 episodes E01-E07 across both discs, extra moved to extras\'
    Assert-Equal 'Joking Apart-S02-Extra01.mp4' (Get-RelNames (Join-Path $ja 'Specials')) 'apply: the Series 2 extra moved to the series-level Specials folder, beside the season folders'
    Assert-True (Test-Path -LiteralPath (Join-Path $ja1 'extras\Joking Apart-Series 1-B1_t05.mp4')) 'apply: the pre-existing extras file is left exactly where it was, under its old name'
    $rows = @(Import-Csv -LiteralPath (Join-Path $ja2 'rename-manifest.csv'))
    Assert-Equal 8 $rows.Count 'one manifest holding a row for every file of both discs'
    Assert-Equal 8 @($rows | Select-Object -ExpandProperty OriginalName -Unique).Count 'manifest rows are distinct (every original name recorded once)'
    Assert-True (Test-Path (Join-Path $ja2 'undo-rename.ps1')) 'undo script copied next to the shared manifest'
    & (Join-Path $ja2 'undo-rename.ps1') *> $null
    Assert-Equal $beforeJa2 (Get-RelNames $ja2) 'undo of the shared manifest restores BOTH discs original names'
    & (Join-Path $ja1 'undo-rename.ps1') *> $null
    Assert-Equal $beforeJa1 (Get-RelNames $ja1) 'undo in Series 1 restores its names and keeps the existing extras file'

    # The real driver in its own process: the disc-token lookup (GroupOf) is passed as a
    # function reference, so it must still work when rename-series.ps1 is run as a script.
    $cliOut = powershell.exe -NoProfile -File (Join-Path $repoRoot 'rename-series.ps1') $ja -NoTmdb *>&1 | Out-String
    Assert-True ($cliOut -notmatch 'CommandNotFoundException|not\s+recognized') 'rename-series.ps1 run as a script: disc-token lookup works (no "not recognized" error)'
    Assert-True ($cliOut -match 'Joking Apart-S02-E07') 'rename-series.ps1 run as a script: Disc 2 file is planned'
    Assert-Equal $beforeJa2 (Get-RelNames $ja2) 'rename-series.ps1 dry run as a script: nothing renamed'

    # -WhatIf is accepted (users expect it, as undo-rename.ps1 has it) and always means dry run.
    $cliOut = powershell.exe -NoProfile -File (Join-Path $repoRoot 'rename-series.ps1') $ja -NoTmdb -Apply -Yes -WhatIf *>&1 | Out-String
    Assert-True ($cliOut -notmatch 'parameter cannot be found') 'rename-series.ps1 accepts -WhatIf'
    Assert-True ($cliOut -match 'DRY RUN') '-WhatIf with -Apply still runs as a dry run'
    Assert-Equal $beforeJa2 (Get-RelNames $ja2) '-WhatIf -Apply -Yes: nothing renamed'
    Assert-True (-not (Test-Path (Join-Path $ja2 'rename-manifest.csv'))) '-WhatIf -Apply -Yes: no manifest written'

    # Two units sharing a folder never plan the same Extra number, and a later run continues
    $x = Join-Path $tempRoot 'Xtra\Xtra-Series 1'
    New-DiscFolder $x @{ 'Xtra-Series 1 Disc 1-A_t00.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t01.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t03.mkv' = 30; 'Xtra-Series 1 Disc 1-A_t02.mkv' = 4; 'Xtra-Series 1 Disc 2-B_t00.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t01.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t03.mkv' = 30; 'Xtra-Series 1 Disc 2-B_t02.mkv' = 4 }
    $s = Invoke-SeriesRetroRename -Root $x -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $xe = @($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Extra' } | ForEach-Object { $_.NewName }) -join ','
    Assert-Equal '..\Specials\Xtra-S01-Extra01.mkv,..\Specials\Xtra-S01-Extra02.mkv' $xe 'dry run: the two discs extras are Extra01 and Extra02, never a second Extra01'
    $xep = @($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Episode' } | ForEach-Object { $_.OriginalName }) -join ','
    Assert-Equal 'Xtra-Series 1 Disc 1-A_t00.mkv,Xtra-Series 1 Disc 1-A_t01.mkv,Xtra-Series 1 Disc 1-A_t03.mkv,Xtra-Series 1 Disc 2-B_t00.mkv,Xtra-Series 1 Disc 2-B_t01.mkv,Xtra-Series 1 Disc 2-B_t03.mkv' $xep 'episodes numbered in disc-then-title order'

    $r = Join-Path $tempRoot 'Resume\Resume-Series 1'
    New-DiscFolder $r @{ 'Resume-S01-E01.mkv' = 30; 'Resume-S01-E02.mkv' = 30; 'Resume-Series 1 Disc 2-B_t00.mkv' = 30 }
    $s = Invoke-SeriesRetroRename -Root $r -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Resume-S01-E03.mkv' (@($s.Results[0].Result.Plan)[0].NewName) 'a Disc 2 file whose Disc 1 is already renamed in the same folder starts at E03'

    # ---------------------------------------------------------------------------
    Write-Host "`nOne classification per folder: a lone token-less extra is judged with the rest (Boys from the Black Stuff)" -ForegroundColor Cyan

    Assert-Equal '2' "$(Get-DiscFolderNumber 'Boys from the Black Stuff-Disc 2')" 'disc folder: "<Show>-Disc 2"'
    Assert-Equal '3' "$(Get-DiscFolderNumber 'Show Disc 3')" 'disc folder: "<Show> Disc 3"'
    Assert-True ($null -eq (Get-DiscFolderNumber 'Disc 2 extras')) 'a folder merely containing "Disc 2" is not a disc folder'

    $boys = Join-Path $tempRoot 'Boys from the Black Stuff'
    New-DiscFolder (Join-Path $boys 'Season 1') @{ 'B1_T00-1.mp4' = 4.4; 'Boys from the Black Stuff-Disc 1 - E01.mp4' = 60 }
    New-DiscFolder (Join-Path $boys 'Boys from the Black Stuff-Disc 2') @{ 'Boys from the Black Stuff-Disc 2 - E01.mp4' = 54; 'Boys from the Black Stuff-Disc 2 - E02.mp4' = 57 }
    New-DiscFolder (Join-Path $boys 'Boys from the Black Stuff-Disc 3') @{ 'Boys from the Black Stuff-Disc 3 - E01.mp4' = 67; 'Boys from the Black Stuff-Disc 3 - E02.mp4' = 67 }

    $units = @(Get-SeriesRenameUnits -Root $boys)
    Assert-Equal 'S1D0,S1D2,S1D3' (($units | ForEach-Object { "S$($_.Season)D$($_.Disc)" }) -join ',') 'Season 1 folder is one unit; "<Show>-Disc N" folders beside it are found and belong to Season 1'

    $s = Invoke-SeriesRetroRename -Root $boys -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 3 $s.Units 'three units: Season 1 plus the two Disc folders'
    $sp = @($s.Results[0].Result.Plan)
    Assert-Equal '..\Specials\Boys from the Black Stuff-S01-Extra01.mp4' (($sp | Where-Object { $_.OriginalName -eq 'B1_T00-1.mp4' }).NewName) 'the 4-minute token-less file is an EXTRA (judged against the Disc 1 title, not alone)'
    Assert-Equal 'Boys from the Black Stuff-S01-E01.mp4' (($sp | Where-Object { $_.OriginalName -like '*Disc 1*' }).NewName) 'the Disc 1 title is E01'
    Assert-Equal 'Boys from the Black Stuff-S01-E02.mp4,Boys from the Black Stuff-S01-E03.mp4' (@($s.Results[1].Result.Plan | ForEach-Object { $_.NewName }) -join ',') '"-Disc 2" folder continues the season at E02'
    Assert-Equal 'Boys from the Black Stuff-S01-E04.mp4,Boys from the Black Stuff-S01-E05.mp4' (@($s.Results[2].Result.Plan | ForEach-Object { $_.NewName }) -join ',') '"-Disc 3" folder continues at E04'

    $null = Invoke-SeriesRetroRename -Root (Join-Path $boys 'Boys from the Black Stuff-Disc 2') -StartEpisode 2 -Apply -Yes -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1
    $s = Invoke-SeriesRetroRename -Root (Join-Path $boys 'Boys from the Black Stuff-Disc 3') -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Boys from the Black Stuff-S01-E04.mp4' (@($s.Results[0].Result.Plan)[0].NewName) '"-Disc 3" pointed at alone continues after the renamed "-Disc 2" folder'

    # ---------------------------------------------------------------------------
    Write-Host "`nNumbers in names sort as numbers (Blackadder: 'S01E (1)' .. 'S01E (12)')" -ForegroundColor Cyan

    $bl = Join-Path $tempRoot 'Blackadder\Season 1'
    $blFiles = @{}
    foreach ($n in 1..12) { $blFiles["Blackadder-S01E ($n).mp4"] = 30 + ($n % 3) }
    New-DiscFolder $bl $blFiles
    $s = Invoke-SeriesRetroRename -Root $bl -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $byOrig = @{}; foreach ($p in $s.Results[0].Result.Plan) { $byOrig[$p.OriginalName] = $p.NewName }
    Assert-Equal 'Blackadder-S01-E02.mp4' $byOrig['Blackadder-S01E (2).mp4'] '"(2)" becomes E02, not E05'
    Assert-Equal 'Blackadder-S01-E10.mp4' $byOrig['Blackadder-S01E (10).mp4'] '"(10)" becomes E10, not E02'
    Assert-Equal 'Blackadder-S01-E12.mp4' $byOrig['Blackadder-S01E (12).mp4'] '"(12)" becomes E12'

    # ---------------------------------------------------------------------------
    Write-Host "`nPlay-all check runs per disc when one folder holds several discs" -ForegroundColor Cyan

    $pa = Join-Path $tempRoot 'PlayAll\Series 1'
    New-DiscFolder $pa @{ 'P-Series 1 Disc 1-t00.mkv' = 30; 'P-Series 1 Disc 1-t01.mkv' = 30; 'P-Series 1 Disc 1-t02.mkv' = 30; 'P-Series 1 Disc 1-t03.mkv' = 90
                          'P-Series 1 Disc 2-t00.mkv' = 30; 'P-Series 1 Disc 2-t01.mkv' = 30; 'P-Series 1 Disc 2-t02.mkv' = 30; 'P-Series 1 Disc 2-t03.mkv' = 90 }
    $s = Invoke-SeriesRetroRename -Root $pa -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $paExtras = @($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Extra' } | ForEach-Object { $_.OriginalName }) -join ','
    Assert-Equal 'P-Series 1 Disc 1-t03.mkv,P-Series 1 Disc 2-t03.mkv' $paExtras 'each disc has its own play-all, and both are extras (neither is the sum of the whole folder)'
    Assert-Equal 6 @($s.Results[0].Result.Plan | Where-Object { $_.Kind -eq 'Episode' }).Count 'the six 30-minute titles are episodes'

    # ---------------------------------------------------------------------------
    Write-Host "`nSpecials (S00): flagged when much too long, made only on request" -ForegroundColor Cyan

    $sp = Join-Path $tempRoot 'Black Stuff'
    $sp1 = Join-Path $sp 'Season 1'
    $sp2 = Join-Path $sp 'Black Stuff-Disc 2'
    $sp3 = Join-Path $sp 'Black Stuff-Disc 3'
    New-DiscFolder $sp1 @{ 'B1_T00-1.mp4' = 4.4; 'Black Stuff-Disc 1 - E01.mp4' = 102 }
    New-DiscFolder $sp2 @{ 'Black Stuff-Disc 2 - E01.mp4' = 54; 'Black Stuff-Disc 2 - E02.mp4' = 57 }
    New-DiscFolder $sp3 @{ 'Black Stuff-Disc 3 - E01.mp4' = 68; 'Black Stuff-Disc 3 - E02.mp4' = 68 }
    $beforeSp1 = Get-RelNames $sp1

    $s = Invoke-SeriesRetroRename -Root $sp1 -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $long = @($s.Results[0].Result.Plan | Where-Object { $_.OriginalName -like '*Disc 1*' })[0]
    Assert-Equal 'Episode' $long.Kind 'a much-too-long title is NOT made a special on its own (a double episode looks the same)'
    Assert-True ($long.Note -match 'special\?') '...but its note asks whether it is a special'

    $s = Invoke-SeriesRetroRename -Root $sp -NoTmdb -MarkKinds @{ '*Disc 1 - E01*' = 'Special' } -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    $p1 = @($s.Results[0].Result.Plan)
    Assert-Equal '..\Specials\Black Stuff-S00-E01.mp4' (($p1 | Where-Object { $_.Kind -eq 'Special' }).NewName) '-MarkKinds Special: <Title>-S00-E01 in the series Specials folder'
    Assert-Equal '..\Specials\Black Stuff-S01-Extra01.mp4' (($p1 | Where-Object { $_.Kind -eq 'Extra' }).NewName) 'the short title is still an extra'
    Assert-Equal 'Black Stuff-S01-E01.mp4,Black Stuff-S01-E02.mp4' (@($s.Results[1].Result.Plan | ForEach-Object { $_.NewName }) -join ',') 'a special uses no episode number: Disc 2 starts at E01'

    # An episode missing from the rip: -StartEpisode on the folder after the gap wins over the suggestion.
    $s = Invoke-SeriesRetroRename -Root $sp3 -NoTmdb -StartEpisode 4 -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Black Stuff-S01-E04.mp4,Black Stuff-S01-E05.mp4' (@($s.Results[0].Result.Plan | ForEach-Object { $_.NewName }) -join ',') '-StartEpisode 4 on a lone Disc 3 wins (E03 was never ripped)'

    # Apply the special, check undo brings it back and retires the manifest.
    $null = Invoke-SeriesRetroRename -Root $sp1 -Apply -Yes -NoTmdb -MarkKinds @{ '*Disc 1 - E01*' = 'Special' } -GetDuration $getDuration -UndoScriptSource $undoPath 6>&1
    Assert-Equal 'Black Stuff-S00-E01.mp4,Black Stuff-S01-Extra01.mp4' (Get-RelNames (Join-Path $sp 'Specials')) 'apply: special and extra both in Specials'
    Assert-Equal 'Special' (@(Import-Csv -LiteralPath (Join-Path $sp1 'rename-manifest.csv') | Where-Object { $_.NewName -like '*S00-E01*' })[0].Kind) 'manifest records Kind Special'
    $s = Invoke-SeriesRetroRename -Root $sp2 -NoTmdb -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 'Black Stuff-S01-E01.mp4' (@($s.Results[0].Result.Plan)[0].NewName) 'a renamed special is not counted as an earlier episode'
    & (Join-Path $sp1 'undo-rename.ps1') *> $null
    Assert-Equal $beforeSp1 (Get-RelNames $sp1) 'undo: special and extra back under their original names'

    # A second special in the same Specials folder takes S00-E02.
    New-Item -ItemType Directory -Path (Join-Path $sp 'Specials') -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $sp 'Specials\Black Stuff-S00-E01.mp4') | Out-Null
    $s = Invoke-SeriesRetroRename -Root $sp1 -NoTmdb -MarkKinds @{ '*Disc 1 - E01*' = 'Special' } -GetDuration $getDuration 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal '..\Specials\Black Stuff-S00-E02.mp4' ((@($s.Results[0].Result.Plan) | Where-Object { $_.Kind -eq 'Special' }).NewName) 'an existing S00-E01 in Specials pushes the new special to S00-E02'
    Remove-Item -LiteralPath (Join-Path $sp 'Specials') -Recurse -Force

    # Edit at the prompt: "e" then "2s" makes row 2 a special.
    $answers = New-Object System.Collections.Queue
    'e', '2s', 'y' | ForEach-Object { $answers.Enqueue($_) }
    $readEdit = { param($p) if ($answers.Count -gt 0) { $answers.Dequeue() } else { $null } }.GetNewClosure()
    $null = Invoke-SeriesRetroRename -Root $sp1 -Apply -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath -ReadInput $readEdit 6>&1
    Assert-True (Test-Path -LiteralPath (Join-Path $sp 'Specials\Black Stuff-S00-E01.mp4')) 'prompt edit "2s": row 2 renamed as special S00-E01'
    & (Join-Path $sp1 'undo-rename.ps1') *> $null

    # ---------------------------------------------------------------------------
    Write-Host "`nNo answer (end of input) at the prompt renames nothing" -ForegroundColor Cyan

    $readEof = { param($p) $null }
    $s = Invoke-SeriesRetroRename -Root $sp2 -Apply -NoTmdb -GetDuration $getDuration -UndoScriptSource $undoPath -ReadInput $readEof 6>&1 | Where-Object { $_ -is [pscustomobject] } | Select-Object -Last 1
    Assert-Equal 1 $s.Declined 'end of input declines the folder'
    Assert-Equal 0 $s.Renamed 'end of input: nothing renamed'
    Assert-True (-not (Test-Path (Join-Path $sp2 'rename-manifest.csv'))) 'end of input: no manifest written'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "Passed: $script:Passed  Failed: $script:Failed" -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
if ($script:Failed -gt 0) { exit 1 }
