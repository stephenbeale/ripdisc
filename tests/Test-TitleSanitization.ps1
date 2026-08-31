#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for the $safeTitle path-sanitization expression used to build every
    output path / filename from -title.

.DESCRIPTION
    Extracts the four `$safeTitle = ...` expressions (two in rip-disc.ps1, two in
    continue-rip.ps1) directly from source text with Select-String, so the tests
    exercise the shipped expression rather than a reimplementation of it, and asserts
    all four are identical before evaluating any of them.

    Covers the "W." bug: a trailing "." or space in a title is silently dropped by
    Windows when the directory/file is actually created on disk, so $safeTitle must
    strip it too or every downstream comparison that trusts $safeTitle as literal text
    (rather than re-reading the real name back from disk) mismatches and mis-renames
    files. Also re-covers the existing illegal-character handling (PR #131) to make
    sure the added TrimEnd didn't regress it.

    No disc, no MakeMKV and no HandBrake are needed.

.EXAMPLE
    .\tests\Test-TitleSanitization.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

# ---------- pull every $safeTitle assignment out of both scripts ----------

function Get-SafeTitleExpressions {
    param([string]$ScriptPath)

    $matches = Select-String -Path $ScriptPath -Pattern '^\s*\$safeTitle\s*=\s*(.+)$'
    if (-not $matches -or $matches.Count -eq 0) {
        throw "No `$safeTitle assignment found in $ScriptPath"
    }
    return $matches | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
}

$ripDiscExprs = Get-SafeTitleExpressions -ScriptPath $ripDiscPath
$continueExprs = Get-SafeTitleExpressions -ScriptPath $continuePath
$allExprs = @($ripDiscExprs) + @($continueExprs)

# ---------- tiny assertion harness ----------

$script:Passed = 0
$script:Failed = 0

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -eq $Expected) {
        $script:Passed++
    } else {
        $script:Failed++
        Write-Host "FAIL: $Message" -ForegroundColor Red
        Write-Host "  expected: '$Expected'" -ForegroundColor Red
        Write-Host "  actual:   '$Actual'" -ForegroundColor Red
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        $script:Passed++
    } else {
        $script:Failed++
        Write-Host "FAIL: $Message" -ForegroundColor Red
    }
}

# ---------- consistency: all four occurrences must be identical ----------

Assert-True ($allExprs.Count -eq 4) "expected 4 `$safeTitle assignments total (2 per script), found $($allExprs.Count)"

$distinct = $allExprs | Select-Object -Unique
Assert-True ($distinct.Count -eq 1) "all `$safeTitle expressions should be identical across both scripts; found $($distinct.Count) distinct variant(s): $($distinct -join ' | ')"

# ---------- evaluate the real expression against sample titles ----------

function Get-SafeTitle {
    param([string]$Title, [string]$Expression)

    # $title is the only free variable the extracted expression references.
    $title = $Title
    return Invoke-Expression $Expression
}

$expr = $distinct[0]

$cases = @(
    @{ Title = 'W.';                              Expected = 'W' }
    @{ Title = 'W..';                              Expected = 'W' }
    @{ Title = 'Inception ';                       Expected = 'Inception' }
    @{ Title = 'Inception. ';                      Expected = 'Inception' }
    @{ Title = 'Southpaw';                         Expected = 'Southpaw' }
    @{ Title = 'The Arena Hawaii 05/06 Highlight Reel - ASL'; Expected = 'The Arena Hawaii 05_06 Highlight Reel - ASL' }
    @{ Title = 'Se7en: The Director''s Cut';       Expected = 'Se7en_ The Director''s Cut' }
    @{ Title = 'M*A*S*H';                          Expected = 'M_A_S_H' }
    @{ Title = 'W';                                Expected = 'W' }
)

foreach ($case in $cases) {
    $actual = Get-SafeTitle -Title $case.Title -Expression $expr
    Assert-Equal $actual $case.Expected "`$safeTitle for -title '$($case.Title)'"
}

# ---------- summary ----------

Write-Host ""
Write-Host "Passed: $script:Passed" -ForegroundColor Green
if ($script:Failed -gt 0) {
    Write-Host "Failed: $script:Failed" -ForegroundColor Red
    exit 1
} else {
    Write-Host "All tests passed." -ForegroundColor Green
}
