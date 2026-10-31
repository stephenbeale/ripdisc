#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the pre-rip existing-folder check: disc titles vs files already in the
    title's folder (suspected duplicates), the table, and the Enter/O/E/A choice.

.DESCRIPTION
    Function bodies are lifted out of rip-disc.ps1 with the AST parser. File lengths are
    injected (no real media needed); the folder scan runs on a temp folder.

.EXAMPLE
    .\tests\Test-ExistingFolderCheck.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
. (Join-Path $repoRoot 'SeriesEpisodes.ps1')

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ScriptPath has parse errors" }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}
foreach ($f in 'Get-FolderVideoFiles', 'Compare-DiscToFolder', 'Format-FileSize', 'Show-DiscFolderComparison', 'Read-ExistingFolderChoice', 'Get-ShortText') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $f)))
}

$script:Passed = 0; $script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because`n        expected: [$Expected]`n        actual:   [$Actual]" -ForegroundColor Red }
}
function Assert-True { param([bool]$Condition, [string]$Because) Assert-Equal $true $Condition $Because }
function New-InputQueue {
    param([string[]]$Answers)
    $q = New-Object System.Collections.Queue
    foreach ($a in $Answers) { $q.Enqueue($a) }
    return { param($p) if ($q.Count -gt 0) { $q.Dequeue() } else { $null } }.GetNewClosure()
}
function New-File { param($Rel, $Secs, $Bytes = 100MB) [pscustomobject]@{ RelPath = $Rel; FullPath = "X:\Film\$Rel"; Seconds = $Secs; Bytes = $Bytes } }

# Real case 2026-10-31: F:\DVDs\Erin Brockovich (Feature + 4 extras, lengths as Windows reads them)
$erin = @(
    (New-File 'Erin Brockovich-Feature.mp4' 7555 1.35GB),
    (New-File 'extras\Erin Brockovich-B1_t01.mp4' 147),
    (New-File 'extras\Erin Brockovich-B2_t02.mp4' 897),
    (New-File 'extras\Erin Brockovich-B3_t03.mp4' 240),
    (New-File 'extras\Erin Brockovich-B4_t04.mp4' 1808))

Write-Host "`nCompare-DiscToFolder" -ForegroundColor Cyan
$c = Compare-DiscToFolder -DiscSeconds @(7556, 897, 240, 1807, 146, 12) -Files $erin
Assert-Equal 5 $c.Matched "same disc again: all 5 titles over 2 min match (1-2 s encode drift allowed)"
Assert-Equal 5 @($c.Rows).Count "the 12-second title is below the 2-minute minimum and not listed"
Assert-Equal 0 @($c.Unmatched).Count "every existing file accounted for"
Assert-Equal 'Erin Brockovich-Feature.mp4' $c.Rows[0].Match.RelPath "longest title matched to the Feature"

$c = Compare-DiscToFolder -DiscSeconds @(7556, 1500, 897) -Files $erin
Assert-Equal 2 $c.Matched "a different disc with one new extra: 2 matched"
Assert-Equal $null ($c.Rows | Where-Object { $_.Seconds -eq 1500 }).Match "the 25:00 title is new"
Assert-Equal 3 @($c.Unmatched).Count "three existing extras are not on this disc"

$c = Compare-DiscToFolder -DiscSeconds @(7400) -Files $erin
Assert-Equal 0 $c.Matched "a 2:03:20 cut is NOT the 2:05:55 feature (beyond 1%) - a different version"

$c = Compare-DiscToFolder -DiscSeconds @(1320, 1320, 1320, 1320) -Files @((New-File 'a.mp4' 1320), (New-File 'b.mp4' 1321))
Assert-Equal 1 @($c.Rows).Count "identical disc lengths collapse into one row"
Assert-Equal 4 $c.Rows[0].Count "with their count"
Assert-Equal 1 @($c.Unmatched).Count "one file per row: the other same-length file stays unmatched"

$c = Compare-DiscToFolder -DiscSeconds @(600) -Files @((New-File 'x.mp4' $null))
Assert-Equal 0 $c.Matched "a file whose length is unknown never matches"
$c = Compare-DiscToFolder -DiscSeconds @() -Files $erin
Assert-Equal 0 @($c.Rows).Count "no disc lengths: no rows"
Assert-Equal 5 @($c.Unmatched).Count "all files still listed"

Write-Host "`nGet-FolderVideoFiles" -ForegroundColor Cyan
$tmp = Join-Path ([IO.Path]::GetTempPath()) "ripdisc-existing-$([guid]::NewGuid().ToString('N'))"
try {
    New-Item -ItemType Directory -Path (Join-Path $tmp 'Film [2000]\extras\deeper') -Force | Out-Null
    $film = Join-Path $tmp 'Film [2000]'
    foreach ($n in 'Film-Feature.mp4', 'folder.jpg', '._Film-Feature.mp4', 'extras\Film-B1.mkv', 'extras\deeper\too-deep.mp4') {
        Set-Content -LiteralPath (Join-Path $film $n) -Value 'x'
    }
    $files = @(Get-FolderVideoFiles -Folder $film -GetSeconds { param($p) 60 })
    Assert-Equal 'Film-Feature.mp4|extras\Film-B1.mkv' (($files | ForEach-Object { $_.RelPath }) -join '|') "videos in the folder and one level down; no images, no ._ files, nothing deeper; brackets in the path are fine"
    Assert-Equal 60 $files[0].Seconds "length from the injected reader"
    Assert-Equal 0 @(Get-FolderVideoFiles -Folder (Join-Path $tmp 'missing')).Count "missing folder: nothing"
} finally { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }

Write-Host "`nShow-DiscFolderComparison" -ForegroundColor Cyan
$c = Compare-DiscToFolder -DiscSeconds @(7556, 1500) -Files $erin
$text = (Show-DiscFolderComparison -Comparison $c -Folder 'F:\DVDs\Erin Brockovich' 6>&1 | Out-String)
Assert-True ($text -match 'ALREADY HAVE\?\s+Erin Brockovich-Feature\.mp4\s+2:05:55\s+1\.35 GB') "matched row shows file, length and size"
Assert-True ($text -match '25:00\s+new') "new row"
Assert-True ($text -match 'In the folder but not on this disc') "unmatched files listed"
Assert-True ($text -match '1 of 2 title\(s\)') "summary line"
$text = (Show-DiscFolderComparison -Comparison (Compare-DiscToFolder -Files $erin) -Folder 'F:\x' -NoDiscLengths 6>&1 | Out-String)
Assert-True ($text -match 'could not be read' -and $text -match 'Files in the folder') "no disc lengths: files-only table"

Write-Host "`nRead-ExistingFolderChoice" -ForegroundColor Cyan
$c = Compare-DiscToFolder -DiscSeconds @(7556, 897) -Files $erin
$opened = New-Object System.Collections.Generic.List[string]
$open = { param($p) $opened.Add($p) }.GetNewClosure()
Assert-Equal 'continue' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @('')) -OpenPath $open 6>$null) "Enter continues"
Assert-Equal 'abort' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @('a')) -OpenPath $open 6>$null) "A aborts"
Assert-Equal 'abort' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @()) -OpenPath $open 6>$null) "end of input aborts, never rips silently"
Assert-Equal 'continue' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @('o', '')) -OpenPath $open 6>$null) "O opens, then asks again"
Assert-Equal 'X:\Film\Erin Brockovich-Feature.mp4|X:\Film\extras\Erin Brockovich-B2_t02.mp4' ($opened -join '|') "O: one Explorer window per folder (film folder, extras), each with its duplicate highlighted"
$opened.Clear()
$many = Compare-DiscToFolder -DiscSeconds (1..8 | ForEach-Object { 600 + $_ * 60 }) -Files (1..8 | ForEach-Object { New-File "e$_.mp4" (600 + $_ * 60) })
$null = Read-ExistingFolderChoice -Comparison $many -Folder 'F:\Many' -ReadInput (New-InputQueue @('O', 'A')) -OpenPath $open 6>$null
Assert-Equal 'X:\Film\e8.mp4' ($opened -join '|') "eight duplicates in one folder: ONE Explorer window, the longest (first-listed) highlighted"
$opened.Clear()
$null = Read-ExistingFolderChoice -Comparison (Compare-DiscToFolder -Files $erin) -Folder 'F:\x' -ReadInput (New-InputQueue @('O', 'A')) -OpenPath $open 6>$null
Assert-Equal 'X:\Film\Erin Brockovich-Feature.mp4|X:\Film\extras\Erin Brockovich-B1_t01.mp4' ($opened -join '|') "no matches (no disc lengths): O shows the folder's files, one window per folder"
Assert-Equal 'Blu-ray' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @('e', ' Blu-ray ')) -OpenPath $open 6>$null) "E returns the edition label, trimmed"
Assert-Equal 'abort' (Read-ExistingFolderChoice -Comparison $c -Folder 'F:\x' -ReadInput (New-InputQueue @('e', '', 'x', 'a')) -OpenPath $open 6>$null) "blank edition goes back; junk re-asks"

Write-Host "`nWiring in rip-disc.ps1" -ForegroundColor Cyan
$src = Get-Content -LiteralPath $ripDiscPath -Raw
$checkAt = $src.IndexOf('EXISTING FOLDER: SUSPECTED DUPLICATES')
$clashAt = $src.IndexOf("NEVER RIP A FILM INTO ANOTHER FILM'S FOLDER")
Assert-True ($checkAt -gt 0 -and $checkAt -lt $clashAt) "the check runs before the folder-clash guard (so the table is seen before Disc 1 is moved to 'Title (Year)')"
Assert-True ($src -match '\$title = "\$title - \$folderChoice"') "an edition becomes part of the title (files, folder, continue command stay consistent)"

Write-Host "`n$($script:Passed) passed, $($script:Failed) failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed -gt 0) { exit 1 }
