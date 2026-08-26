#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for Get-ContinueRipCommand, the failure-time continue-rip.ps1 retry suggestion.

.DESCRIPTION
    Runs against the real function body lifted out of rip-disc.ps1 at run time via the
    PowerShell AST parser, so the tests cannot drift from the shipped code without failing.
    No disc, MakeMKV or HandBrake needed.

.EXAMPLE
    .\tests\Test-ContinueRipCommand.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

if (-not (Test-Path $ripDiscPath)) { throw "Cannot find $ripDiscPath - run this from inside the repo." }

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }

    $fn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1

    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-ContinueRipCommand')))

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
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}

function Assert-Null {
    param($Actual, [string]$Because)

    if ($null -eq $Actual) {
        $script:Passed++
        Write-Host "  PASS  $Because" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
        Write-Host "        expected: [`$null]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}

Write-Host "`nStep-to-FromStep mapping" -ForegroundColor Cyan

Assert-Equal '.\continue-rip.ps1 -title "Powers of Three" -FromStep handbrake -Documentary' `
    (Get-ContinueRipCommand -Title 'Powers of Three' -RemainingStepNumbers @(2, 3, 4) -Documentary -Disc 1 -StartEpisode 1 -EpisodeNames @()) `
    'remaining steps 2/3/4 map to -FromStep handbrake (earliest incomplete step wins)'

Assert-Equal '.\continue-rip.ps1 -title "The Matrix" -FromStep organize' `
    (Get-ContinueRipCommand -Title 'The Matrix' -RemainingStepNumbers @(3, 4) -Disc 1 -StartEpisode 1 -EpisodeNames @()) `
    'remaining steps 3/4 map to -FromStep organize'

Assert-Equal '.\continue-rip.ps1 -title "The Matrix" -FromStep open' `
    (Get-ContinueRipCommand -Title 'The Matrix' -RemainingStepNumbers @(4) -Disc 1 -StartEpisode 1 -EpisodeNames @()) `
    'remaining step 4 alone maps to -FromStep open'

Write-Host "`nStep 1 unresolved - nothing to resume" -ForegroundColor Cyan

Assert-Null (Get-ContinueRipCommand -Title 'X' -RemainingStepNumbers @(1, 2, 3, 4) -Disc 1 -StartEpisode 1 -EpisodeNames @()) `
    'returns $null when Step 1 (the MakeMKV rip) is still incomplete - continue-rip.ps1 has nothing to resume from'

Assert-Null (Get-ContinueRipCommand -Title 'X' -RemainingStepNumbers @() -Disc 1 -StartEpisode 1 -EpisodeNames @()) `
    'returns $null when no steps remain at all'

Write-Host "`nDefault values are omitted to keep the command short" -ForegroundColor Cyan

$plain = Get-ContinueRipCommand -Title 'The Matrix' -RemainingStepNumbers @(2) -Season 0 -Disc 1 -StartEpisode 1 -EpisodeNames @()
Assert-Equal '.\continue-rip.ps1 -title "The Matrix" -FromStep handbrake' $plain `
    'Season 0 / Disc 1 / StartEpisode 1 / no OutputDrive match continue-rip.ps1 defaults and are all omitted'

Write-Host "`nNon-default values are carried over" -ForegroundColor Cyan

$series = Get-ContinueRipCommand -Title 'Breaking Bad' -RemainingStepNumbers @(3, 4) -Series -Season 2 -Disc 3 -OutputDrive 'F:' -StartEpisode 5 -EpisodeNames @()
Assert-Equal '.\continue-rip.ps1 -title "Breaking Bad" -FromStep organize -Series -Season 2 -Disc 3 -OutputDrive F: -StartEpisode 5' $series `
    'Series/Season/Disc/OutputDrive/StartEpisode all appear when non-default'

$genre = Get-ContinueRipCommand -Title 'Planet Earth' -RemainingStepNumbers @(2) -Documentary -Extras -Bluray -Disc 1 -StartEpisode 1 -EpisodeNames @()
Assert-Equal '.\continue-rip.ps1 -title "Planet Earth" -FromStep handbrake -Extras -Bluray -Documentary' $genre `
    'switch flags (Extras, Bluray, Documentary) appear in declaration order when set'

$noSound = Get-ContinueRipCommand -Title 'The Matrix' -RemainingStepNumbers @(4) -Disc 1 -StartEpisode 1 -EpisodeNames @() -NoSound
Assert-Equal '.\continue-rip.ps1 -title "The Matrix" -FromStep open -NoSound' $noSound `
    '-NoSound is carried over when set (continue-rip.ps1 accepts it and it is functional there)'

Write-Host "`nQuoting" -ForegroundColor Cyan

$quotedTitle = Get-ContinueRipCommand -Title 'The "Best" Movie' -RemainingStepNumbers @(4) -Disc 1 -StartEpisode 1 -EpisodeNames @()
Assert-Equal '.\continue-rip.ps1 -title "The `"Best`" Movie" -FromStep open' $quotedTitle `
    'a literal double quote in the title is escaped so it still parses as one argument when pasted'

$episodeNames = Get-ContinueRipCommand -Title 'The Blues' -RemainingStepNumbers @(3) -Documentary -Series -Disc 3 -StartEpisode 1 -EpisodeNames @('The Road to Memphis', 'Warming by the Devil''s Fire')
Assert-Equal '.\continue-rip.ps1 -title "The Blues" -FromStep organize -Series -Disc 3 -Documentary -EpisodeNames "The Road to Memphis", "Warming by the Devil''s Fire"' $episodeNames `
    '-EpisodeNames is carried over, comma-joined and individually quoted'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
