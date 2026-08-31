#Requires -Version 5.1
<#
.SYNOPSIS
    Confirms Series/Extras file-prefixing derives its prefix from the real on-disk
    directory name instead of trusting $safeTitle as text.

.DESCRIPTION
    Two things are checked:

    1. A live filesystem test (not a text/logic test - this repo's prior sanitization
       work only ever reasoned about Windows' trailing-dot/space stripping from a Linux
       sandbox, which doesn't reproduce it) confirming, on this actual Windows machine,
       that creating a directory named e.g. "W." really is created on disk as "W", and
       that Get-Item reflects the trimmed name - the assumption the whole "read the name
       back from disk" fix depends on. Also confirms Get-Item's .Parent.Name correctly
       identifies a title folder from its "extras" subfolder, matching the shape used by
       the Extras-mode fix (a title folder that's already-trimmed by the time it's
       created, since Get-SafeTitle now trims before $finalOutputDir is ever built - see
       below for why an *untrimmed* intermediate segment is deliberately not tested).

       Investigated but deliberately NOT asserted here: trailing-dot/space trimming only
       applies to the FINAL segment of a path being resolved - an untrimmed
       *intermediate* segment (e.g. resolving ".../Title. /extras" in one call) is NOT
       auto-trimmed by Get-Item/Test-Path, and fails to resolve at all. This doesn't
       affect the shipped code: Get-SafeTitle trims $safeTitle before any path
       ($finalOutputDir included) is ever built from it, so an untrimmed intermediate
       segment is never actually constructed. Noted here so the reasoning isn't lost.

    2. A source-inspection regression guard confirming the Series-mode and Extras-mode
       prefix blocks in both scripts build their prefix via Get-Item (reading the real
       on-disk name) rather than interpolating $safeTitle directly - so a future edit
       that reintroduces the text-based prefix (and with it, the whole "W."-style bug
       class Get-SafeTitle's TrimEnd only partially covers) fails this test.

    Movie mode's two prefix branches already used Get-Item before this fix (see the
    2026-08-24 CLAUDE.md session notes) and are included here only as a sanity check
    that the pattern they set is what Series/Extras now follow too.

.EXAMPLE
    .\tests\Test-PrefixDerivedFromDisk.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
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
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
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

# ---------- live filesystem behaviour ----------

Write-Host "`nLive Windows behaviour: trailing dot/space really is stripped on disk" -ForegroundColor Cyan

$tempRoot = Join-Path $env:TEMP "ripdisc-tests-prefix-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    # The bug's own trigger case: requesting "W." as a directory's own (final-segment)
    # name. This is the shape every real New-Item/mkdir call in the script actually
    # makes - $safeTitle is always the last segment of ITS OWN creation call, even when
    # a deeper path (e.g. "...\extras") gets created in the same -Force invocation,
    # because Windows creates one path segment at a time under the hood.
    $dotDirRequested = Join-Path $tempRoot "W."
    New-Item -ItemType Directory -Path $dotDirRequested -Force | Out-Null
    Assert-Equal 'W' (Get-Item -LiteralPath $dotDirRequested).Name 'a directory requested as "W." is actually created on disk as "W", and Get-Item -LiteralPath on the original (untrimmed) request string still resolves it (the "W." by Oliver Stone case)'

    # Get-Item's .Parent.Name (the shape used for the Extras-mode fix, where the real
    # title folder is the extras folder's parent). $safeTitle is already trimmed by
    # Get-SafeTitle before $finalOutputDir is ever built (see rip-disc.ps1/continue-rip.ps1
    # - $safeTitle is computed before any path construction), so production never asks
    # Windows to create an untrimmed *intermediate* path segment - this mirrors that by
    # creating the title folder pre-trimmed, exactly as Get-SafeTitle's output would be.
    $titleDir = Join-Path $tempRoot "W"
    $extrasDir = Join-Path $titleDir "extras"
    New-Item -ItemType Directory -Path $extrasDir -Force | Out-Null
    Assert-Equal 'W' (Get-Item -LiteralPath $extrasDir).Parent.Name 'Parent.Name of an "extras" subfolder correctly identifies the title-folder name'

    # A directory with no trimming needed - the read-back should just match
    $plainDir = Join-Path $tempRoot "The Matrix"
    New-Item -ItemType Directory -Path $plainDir -Force | Out-Null
    Assert-Equal 'The Matrix' (Get-Item -LiteralPath $plainDir).Name 'a directory with no trailing dot/space round-trips unchanged'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------- source-inspection regression guard ----------

Write-Host "`nSeries/Extras prefix blocks derive from disk, not from `$safeTitle text" -ForegroundColor Cyan

foreach ($scriptPath in @($ripDiscPath, $continuePath)) {
    $scriptName = Split-Path $scriptPath -Leaf
    $lines = Get-Content -Path $scriptPath

    $seriesBlockStart = ($lines | Select-String -Pattern 'SERIES MODE: Prefix files with title \+ season-disc tag' | Select-Object -First 1).LineNumber
    $extrasBlockStart = ($lines | Select-String -Pattern 'Extras disc: prefix with title only' | Select-Object -First 1).LineNumber
    Assert-True ($null -ne $seriesBlockStart) "Series-mode prefix block found in $scriptName"
    Assert-True ($null -ne $extrasBlockStart) "Extras-mode prefix block found in $scriptName"

    if ($seriesBlockStart) {
        # The $prefix assignment is a handful of lines below the block header
        $seriesWindow = $lines[($seriesBlockStart - 1)..($seriesBlockStart + 14)] -join "`n"
        Assert-True ($seriesWindow -match '\(Get-Item\s+\$seriesBaseDir\)\.Name') "Series-mode prefix in $scriptName is built from (Get-Item `$seriesBaseDir).Name"
        Assert-True ($seriesWindow -notmatch '\$prefix\s*=\s*"\$safeTitle-') "Series-mode prefix in $scriptName no longer interpolates `$safeTitle directly (regression guard)"
    }

    if ($extrasBlockStart) {
        $extrasWindow = $lines[($extrasBlockStart - 1)..($extrasBlockStart + 14)] -join "`n"
        Assert-True ($extrasWindow -match '\(Get-Item\s+\$finalOutputDir\)\.Parent\.Name') "Extras-mode prefix in $scriptName is built from (Get-Item `$finalOutputDir).Parent.Name"
        Assert-True ($extrasWindow -notmatch '-notlike\s*\("\$safeTitle-\*"\)') "Extras-mode prefix in $scriptName no longer matches against `$safeTitle text directly (regression guard)"
    }
}

# ---------- summary ----------

Write-Host ""
Write-Host "$script:Passed/$($script:Passed + $script:Failed) passed" -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
if ($script:Failed -gt 0) { exit 1 }
