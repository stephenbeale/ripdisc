#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for spelling-tolerant title matching: a misspelled disc label or -title still finds
    the right pick-list row and the existing folder, and sequels are never confused.

.DESCRIPTION
    Trigger (2026-10-31): the Erin Brockovich DVD's label is "Erin Brokovich". The label
    stayed [1], the year prompt searched IMDb again, and the rip went to a new misspelled
    folder "Erin Brokovich (2000)", bypassing F:\DVDs\Erin Brockovich and its duplicate check.
    Sequel cases come from an audit of all 1181 folders on F: (DVDs, Bluray, Documentaries,
    Surf): the first draft rule paired 63 folders, mostly sequels; the final rule pairs 3,
    all genuinely the same title.

.EXAMPLE
    .\tests\Test-NearMissTitle.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$configPath = Join-Path $repoRoot 'Load-Config.ps1'

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ScriptPath has parse errors" }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}
foreach ($f in 'Get-EditDistance', 'ConvertTo-LooseTitleKey', 'Test-NearMissTitle', 'Find-NearMissMediaFolders') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $configPath -FunctionName $f)))
}
foreach ($f in 'Read-NearMissFolderChoice', 'ConvertTo-TitleKey', 'Test-GenericDiscLabel', 'Clean-DiscName', 'New-TitleCandidate', 'Get-DiscTitleCandidates', 'Get-YearLookup') {
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

Write-Host "`nTest-NearMissTitle - same title, spelled differently" -ForegroundColor Cyan
foreach ($pair in @(
        @('Erin Brokovich', 'Erin Brockovich', 'the real disc label'),
        @('ERIN_BROKOVICH', 'Erin Brockovich', 'case and underscores ignored'),
        @('Hitchiker''s Guide to the Galaxy', 'The Hitchhiker''s Guide to the Galaxy', 'F: pair: missing h, leading The'),
        @('Bridge on the River Kwai', 'The Bridge on the River Kwai', 'F: pair: leading The'),
        @('Ice Age 3 - Dawn of the Dinosaurs', 'Ice Age 3 Dawn of the Dinosaurs', 'F: pair: punctuation'),
        @('Gladiater', 'Gladiator', 'one letter in a 9-letter title'))) {
    Assert-True (Test-NearMissTitle $pair[0] $pair[1]) "$($pair[0]) ~ $($pair[1]) ($($pair[2]))"
}

Write-Host "`nTest-NearMissTitle - different titles (from the F: audit)" -ForegroundColor Cyan
foreach ($pair in @(
        @('Rocky 2', 'Rocky 3'), @('Rocky', 'Rocky 2'), @('Die Hard 2', 'Die Hard 4.0'), @('Iron Man', 'Iron Man 2'),
        @('Spider-Man 2', 'Spiderman'), @('The Godfather Part 2', 'The Godfather Part 3'),
        @('411VM - Issue 14', '411VM - Issue 17'), @('More Mantovani Magic-Disc 1', 'More Mantovani Magic-Disc 2'),
        @('Dumb and Dumber', 'Dumb and Dumber To'), @('Pale Rider', 'Whale Rider'), @('Heat', 'Head'),
        @('Night at the Museum 2-extras', 'Night at the Museum-extras'), @('Erin Brockovich', 'Final Justice with Erin Brockovich'))) {
    Assert-True (-not (Test-NearMissTitle $pair[0] $pair[1])) "$($pair[0]) is not $($pair[1])"
}
Assert-True (-not (Test-NearMissTitle '' 'Rocky')) "empty title never matches"

Write-Host "`nFind-NearMissMediaFolders" -ForegroundColor Cyan
$tmp = Join-Path ([IO.Path]::GetTempPath()) "ripdisc-nearmiss-$([guid]::NewGuid().ToString('N'))"
try {
    foreach ($d in 'Erin Brockovich', 'Rocky 2', 'Rocky 3', 'Gladiator (2000)') { New-Item -ItemType Directory -Path (Join-Path $tmp $d) -Force | Out-Null }
    Assert-Equal 'Erin Brockovich' ((Find-NearMissMediaFolders -ParentDir $tmp -SafeTitle 'Erin Brokovich') -join '|') "the real case: finds the correctly spelled folder"
    Assert-Equal 'Gladiator (2000)' ((Find-NearMissMediaFolders -ParentDir $tmp -SafeTitle 'Gladiater') -join '|') "a folder with a year matches on its name without the year"
    Assert-Equal '' ((Find-NearMissMediaFolders -ParentDir $tmp -SafeTitle 'Erin Brockovich') -join '|') "an exact match is not a near miss (Find-ExistingMediaFolder handles it)"
    Assert-Equal '' ((Find-NearMissMediaFolders -ParentDir $tmp -SafeTitle 'Rocky 4') -join '|') "a new sequel is not offered an earlier film's folder"
    Assert-Equal 0 @(Find-NearMissMediaFolders -ParentDir (Join-Path $tmp 'nope') -SafeTitle 'x').Count "missing parent: nothing"
} finally { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }

Write-Host "`nRead-NearMissFolderChoice" -ForegroundColor Cyan
Assert-Equal 'Erin Brockovich' (Read-NearMissFolderChoice -Title 'Erin Brokovich' -ParentDir 'F:\DVDs' -Folders @('Erin Brockovich') -ReadInput (New-InputQueue @('')) 6>$null) "Enter uses the existing folder"
Assert-Equal $null (Read-NearMissFolderChoice -Title 'Erin Brokovich' -ParentDir 'F:\DVDs' -Folders @('Erin Brockovich') -ReadInput (New-InputQueue @('n')) 6>$null) "N keeps the typed title"
Assert-Equal $null (Read-NearMissFolderChoice -Title 'Erin Brokovich' -ParentDir 'F:\DVDs' -Folders @('Erin Brockovich') -ReadInput (New-InputQueue @()) 6>$null) "end of input keeps the typed title"
Assert-Equal 'B' (Read-NearMissFolderChoice -Title 'x' -ParentDir 'F:\' -Folders @('A', 'B') -ReadInput (New-InputQueue @('7', '2')) 6>$null) "a number picks among several; out of range re-asks"

Write-Host "`nPick list: the correctly spelled film goes above a misspelled label" -ForegroundColor Cyan
$c = Get-DiscTitleCandidates -DiscLabel 'Erin Brokovich' -ImdbCandidates @(
    [pscustomobject]@{ Title = 'Erin Brockovich'; Year = '2000'; MediaType = 'movie'; Kind = 'Movie'; TmdbId = $null; ImdbId = 'tt0195685'; Season = 0; Source = 'IMDb' },
    [pscustomobject]@{ Title = 'Final Justice with Erin Brockovich'; Year = '2003'; MediaType = 'tv'; Kind = 'TV series'; TmdbId = $null; ImdbId = 'tt0360261'; Season = 0; Source = 'IMDb' })
Assert-Equal 'Erin Brockovich|2000' "$($c[0].Title)|$($c[0].Year)" "[1] is Erin Brockovich (2000) - Enter takes the right spelling and year"
Assert-Equal 'disc label' $c[1].Source "the misspelled label is [2]"
$c = Get-DiscTitleCandidates -DiscLabel 'Rocky 4' -ImdbCandidates @(
    [pscustomobject]@{ Title = 'Rocky 3'; Year = '1982'; MediaType = 'movie'; Kind = 'Movie'; TmdbId = $null; ImdbId = 'tt0084602'; Season = 0; Source = 'IMDb' })
Assert-Equal 'disc label' $c[0].Source "another sequel never goes above the label"

Write-Host "`nYear prompt reuses the discovery search" -ForegroundColor Cyan
$script:Searches = 0
function Test-OnlineOnce { $true }
function Get-TMDbTitleCandidates { @() }
function Get-ImdbTitleCandidates { param($SearchTitle) $script:Searches++; @([pscustomobject]@{ Title = $SearchTitle; Year = '2000'; MediaType = 'movie'; Kind = 'Movie'; TmdbId = $null; ImdbId = 'tt1'; Season = 0; Source = 'IMDb' }) }
$script:YearLookupCache = @{ 'erin brokovich' = @([pscustomobject]@{ Title = 'Erin Brockovich'; Year = '2000' }) }
$lookup = Get-YearLookup
$r = @(& $lookup 'Erin Brokovich')
Assert-Equal 0 $script:Searches "seeded by discovery: no second IMDb search"
Assert-Equal 'Erin Brockovich' $r[0].Title "the cached results are returned"
$null = & $lookup 'Erin Brockovich 2000'
$null = & $lookup 'erin brockovich 2000 '
Assert-Equal 1 $script:Searches "[R] with new text searches once; the same text again (any case/spacing) is cached"
$src = Get-Content -LiteralPath $ripDiscPath -Raw
Assert-True ($src -match 'YearLookupCache\["\$searchTitle"\.Trim\(\)\.ToLowerInvariant\(\)\]') "discovery seeds the cache with its own search text"
Assert-True ($src.IndexOf('Find-NearMissMediaFolders -ParentDir $movieParentDir') -lt $src.IndexOf('EXISTING FOLDER: SUSPECTED DUPLICATES')) "the near-miss folder prompt runs before the duplicate check"

Write-Host "`n$($script:Passed) passed, $($script:Failed) failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed -gt 0) { exit 1 }
