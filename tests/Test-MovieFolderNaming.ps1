#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for movie / genre-film folder naming ("Slow Dance" or "Slow Dance (2014)"):
    the shared helpers in Load-Config.ps1, the year prompt in rip-disc.ps1, the -Year
    hand-off to continue-rip.ps1, and that both scripts build movie folders from
    $script:MediaFolderName.

.DESCRIPTION
    Lifts the real functions out of the scripts with the AST parser, so nothing here is
    a reimplementation. Filesystem cases use a temp directory removed afterwards. No
    network, disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-MovieFolderNaming.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$configPath = Join-Path $repoRoot 'Load-Config.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'
$setupPath = Join-Path $repoRoot 'setup.ps1'

foreach ($p in @($configPath, $ripDiscPath, $continuePath, $setupPath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

function Get-ScriptAst {
    param([string]$ScriptPath)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }
    return $ast
}

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $fn = (Get-ScriptAst $ScriptPath).FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

function Get-ScriptParameterNames {
    param([string]$ScriptPath)
    return @((Get-ScriptAst $ScriptPath).ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}

foreach ($name in @('Get-MovieParentFolder', 'Get-MediaFolderName', 'Find-ExistingMediaFolder')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
}
foreach ($name in @('New-TitleCandidate', 'Get-TitleCandidateUrl', 'Format-TitleCandidate', 'Show-TitleCandidates',
                    'Read-FolderYear', 'Get-ContinueRipCommand')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $ripDiscPath $name)))
}

function New-InputQueue {
    param([object[]]$Answers)
    $queue = New-Object System.Collections.Queue
    foreach ($a in $Answers) { $queue.Enqueue($a) }
    return { param($prompt) if ($queue.Count -gt 0) { $queue.Dequeue() } else { $null } }.GetNewClosure()
}

$script:Passed = 0
$script:Failed = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) {
        $script:Passed++
        Write-Host "  PASS  $Because" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
        Write-Host "        expected: $Expected" -ForegroundColor Red
        Write-Host "        actual:   $Actual" -ForegroundColor Red
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Because)
    if ($Condition) {
        $script:Passed++
        Write-Host "  PASS  $Because" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
    }
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-folder-naming-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nFolder name and parent folder" -ForegroundColor Cyan

    Assert-Equal 'Slow Dance (2014)' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '2014') 'year appended in brackets'
    Assert-Equal 'Slow Dance' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '') 'no year, no change'
    Assert-Equal 'Slow Dance' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '14') 'a non-4-digit year is ignored'
    Assert-Equal 'Surf' (Get-MovieParentFolder -Surf) '-Surf films live under Surf'
    Assert-Equal 'Documentaries' (Get-MovieParentFolder -Documentary -Bluray) 'a genre wins over -Bluray (same order as the scripts)'
    Assert-Equal 'Bluray' (Get-MovieParentFolder -Bluray) 'Blu-ray movies under Bluray'
    Assert-Equal 'DVDs' (Get-MovieParentFolder) 'plain movies under DVDs'
    Assert-True ('Slow Dance (2014)-title_t00.mp4' -like ('Slow Dance (2014)' + '-*')) 'parentheses are not wildcards: the movie-mode prefix check still matches'

    # ---------------------------------------------------------------------------
    Write-Host "`nReusing an existing folder" -ForegroundColor Cyan

    $surf = Join-Path $tempRoot 'Surf'
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance')) 'parent folder missing: nothing found'
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance (2014)') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance 2') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dancers (2010)') -Force | Out-Null
    Assert-Equal 'Slow Dance (2014)' (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance') 'Disc 2 finds the "(2014)" folder Disc 1 made; similar names ignored'
    Assert-Equal 'Slow Dance (2014)' (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'slow dance') 'match ignores case, like Windows'
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance') -Force | Out-Null
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance')) 'two candidates (plain and dated): no guess'
    $dvds = Join-Path $tempRoot 'DVDs'
    New-Item -ItemType Directory -Path (Join-Path $dvds 'W_ (2008)') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dvds 'Heat (Director''s Cut)') -Force | Out-Null
    Assert-Equal 'W_ (2008)' (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'W_') 'sanitized title matched literally'
    Assert-Equal 'Heat (Director''s Cut)' (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'Heat (Director''s Cut)') 'regex characters in the title are escaped'
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'Heat')) '"Heat (Director''s Cut)" is not taken as a year folder for "Heat"'

    # ---------------------------------------------------------------------------
    Write-Host "`nYear prompt" -ForegroundColor Cyan

    $hits = @(
        (New-TitleCandidate -Title 'Slow Dance' -Year '2014' -MediaType 'movie' -Kind 'Short' -ImdbId 'tt3453042' -Source 'IMDb'),
        (New-TitleCandidate -Title 'Slow Dance' -Source 'disc label'),
        (New-TitleCandidate -Title 'Slow Dance' -Year '2005' -MediaType 'tv' -Kind 'TV mini-series' -ImdbId 'tt0481050' -Source 'IMDb'))
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('')) 6>$null)) 'Enter: no year'
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('s')) 6>$null)) 'S: no year'
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @()) 6>$null)) 'end of input: no year (the rip is not aborted)'
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('1')) 6>$null) 'row 1 gives its year'
    Assert-Equal '2005' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('2')) 6>$null) 'rows without a year are left out, so row 2 is the 2005 hit'
    Assert-Equal '2013' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('2013')) 6>$null) 'a typed year is used as-is'
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('9', 'x', '1')) 6>$null) 'invalid answers re-ask'
    Assert-Equal '1999' (Read-FolderYear -Title 'Slow Dance' -Candidates @() -ReadInput (New-InputQueue @('1999')) 6>$null) 'no hits: a year can still be typed'
    $script:Searched = @()
    $searchStub = { param($q) $script:Searched += $q; @(New-TitleCandidate -Title 'Slow Dance' -Year '2014' -Kind 'Short' -Source 'IMDb') }
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates @() -Search $searchStub -ReadInput (New-InputQueue @('r', 'slow dance surf', '1')) 6>$null) 'R searches again and the new rows can be picked'
    Assert-Equal 'slow dance surf' ($script:Searched -join ';') 'the typed search text is used'

    # ---------------------------------------------------------------------------
    Write-Host "`nHand-off to continue-rip.ps1" -ForegroundColor Cyan

    Assert-True ((Get-ScriptParameterNames $ripDiscPath) -contains 'Year') 'rip-disc.ps1 has -Year'
    Assert-True ((Get-ScriptParameterNames $continuePath) -contains 'Year') 'continue-rip.ps1 has -Year'
    $cmd = Get-ContinueRipCommand -Title 'Slow Dance' -RemainingStepNumbers @(2) -Surf -Disc 1 -StartEpisode 1 -EpisodeNames @() -Year '2014'
    Assert-Equal '.\continue-rip.ps1 -title "Slow Dance" -FromStep handbrake -Surf -Year 2014' $cmd 'retry command carries the folder year'
    $cmd = Get-ContinueRipCommand -Title 'Slow Dance' -RemainingStepNumbers @(2) -Surf -Disc 1 -StartEpisode 1 -EpisodeNames @()
    Assert-Equal '.\continue-rip.ps1 -title "Slow Dance" -FromStep handbrake -Surf' $cmd 'no year, no -Year'
    $cmd = Get-ContinueRipCommand -Title 'Fargo' -RemainingStepNumbers @(3) -Series -Season 1 -Disc 1 -StartEpisode 1 -EpisodeNames @() -Year '2014'
    Assert-True ($cmd -notmatch '-Year') 'series commands never carry -Year (series folders are unchanged)'

    # ---------------------------------------------------------------------------
    Write-Host "`nBoth scripts build movie folders from the chosen name" -ForegroundColor Cyan

    foreach ($path in @($ripDiscPath, $continuePath)) {
        $text = Get-Content -LiteralPath $path -Raw
        $leaf = Split-Path $path -Leaf
        foreach ($folder in @('Documentaries', 'Tutorials', 'Fitness', 'Music', 'Surf', 'Bluray', 'DVDs')) {
            Assert-True ($text.Contains("`$finalOutputDir = `"`$outputDriveLetter\$folder\`$(`$script:MediaFolderName)`"")) "$leaf`: $folder folder uses the chosen name"
        }
        Assert-True ($text.Contains('$seriesBaseDir = "$outputDriveLetter\Series\$safeTitle"')) "$leaf`: series folder still the plain title"
    }

    # ---------------------------------------------------------------------------
    Write-Host "`nsetup.ps1 and config" -ForegroundColor Cyan

    $setupText = Get-Content -LiteralPath $setupPath -Raw
    Assert-True ($setupText -match 'movieFolderNaming = \$movieFolderNaming') 'setup.ps1 saves movieFolderNaming'
    Assert-True ($setupText -match '\[6/6\] Movie folder naming' -and $setupText -notmatch '/5\]') 'setup steps renumbered to 6'
    $sample = Get-Content -LiteralPath (Join-Path $repoRoot 'ripdisc-config.sample.json') -Raw | ConvertFrom-Json
    Assert-Equal 'title' $sample.movieFolderNaming 'sample config defaults to plain titles'
    $configText = Get-Content -LiteralPath $configPath -Raw
    Assert-True ($configText -match "-in @\('title', 'title-year'\)") 'Load-Config.ps1 accepts only the two styles (anything else falls back to title)'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
