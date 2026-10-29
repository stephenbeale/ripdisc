#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for RipStats.ps1: session-log parsing, the "Estimated finish" line and
    outlier detection, run against fixture logs written to a temp folder.

.EXAMPLE
    .\tests\Test-RipStats.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'RipStats.ps1')

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
    if ($Condition) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because" -ForegroundColor Red }
}

# A ripdisc session log. $Minutes = (rip, encode, idle-after-encode); $Complete $false
# leaves it as a failed session.
function New-VideoLog {
    param([string]$Dir, [string]$Title, [string]$Type, [datetime]$Start, [int[]]$Minutes,
        [int]$Files = 1, [double]$GB = 5, [bool]$Complete = $true, [string[]]$Extra = @())
    $t = $Start
    $f = { param($at, $msg) "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f $at, $msg }
    $lines = @(
        (& $f $t '========== RIP SESSION STARTED =========='), (& $f $t "Title: $Title"), (& $f $t "Type: $Type"),
        (& $f $t 'Disc: 1'), (& $f $t 'Drive: d:'), (& $f $t 'STEP 1/4: Starting MakeMKV rip...'))
    $lines += $Extra | ForEach-Object { & $f $t $_ }
    if (-not $Complete) {
        $lines += & $f $t.AddMinutes(3) '========== ERROR =========='
        $lines += & $f $t.AddMinutes(3) 'Failed at: STEP 1/4: MakeMKV rip'
    } else {
        $t = $t.AddMinutes($Minutes[0])
        $lines += & $f $t "STEP 1/4: MakeMKV rip complete - $Files file(s)"
        for ($i = 0; $i -lt $Files; $i++) { $lines += & $f $t ("  Ripped: A1_t{0:D2}.mkv ({1} GB)" -f $i, ($GB / $Files)) }
        $lines += & $f $t 'STEP 2/4: Starting HandBrake encoding...'
        $t = $t.AddMinutes($Minutes[1])
        $lines += & $f $t "STEP 2/4: HandBrake encoding complete - $Files file(s) encoded"
        $t = $t.AddMinutes($Minutes[2])
        $lines += & $f $t '========== RIP SESSION COMPLETE =========='
        $lines += & $f $t 'Total size: 1.5 GB'
    }
    $name = "{0}_disc1_{1:yyyyMMdd_HHmmss}.log" -f $Title, $Start
    Set-Content -LiteralPath (Join-Path $Dir $name) -Value $lines -Encoding UTF8
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("ripdisc-ripstats-" + [guid]::NewGuid().ToString('N'))
$logDir = Join-Path $tempRoot 'logs'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null

try {
    $base = [datetime]'2026-08-01 10:00:00'
    # Five ordinary movies (40-60 min), one waiting a long time at the end, one documentary
    $i = 0
    foreach ($m in @(@(20, 20, 0), @(25, 20, 0), @(25, 25, 0), @(30, 25, 0), @(30, 30, 0))) {
        New-VideoLog -Dir $logDir -Title "Film $i" -Type 'Movie' -Start $base.AddDays($i) -Minutes $m
        $i++
    }
    New-VideoLog -Dir $logDir -Title 'Slow Film' -Type 'Movie' -Start $base.AddDays(10) -Minutes @(30, 25, 180) `
        -Extra @('MakeMKV killed after 30 consecutive read errors at offset 123 - 1 file(s) salvaged')
    New-VideoLog -Dir $logDir -Title 'Slow Film' -Type 'Movie' -Start $base.AddDays(9) -Complete $false
    New-VideoLog -Dir $logDir -Title 'Some Doc' -Type 'Documentary' -Start $base.AddDays(20) -Minutes @(10, 20, 0)

    # ---------------------------------------------------------------------------
    Write-Host "`nParsing one session log" -ForegroundColor Cyan
    $one = Get-RipLogSummary -Path (Get-ChildItem $logDir -Filter 'Film 0_*').FullName
    Assert-Equal 'Video' $one.Kind 'a ripdisc log is Video'
    Assert-Equal 'Movie' $one.Type 'Type read from the Type: line'
    Assert-Equal 'Film 0' $one.Title 'Title read from the Title: line'
    Assert-Equal 40.0 $one.TotalMinutes 'total = session start to RIP SESSION COMPLETE'
    Assert-Equal 20.0 $one.RipMinutes 'MakeMKV phase = STEP 1 start to complete'
    Assert-Equal 20.0 $one.EncodeMinutes 'HandBrake phase = STEP 2 start to complete'
    Assert-Equal 1 $one.Files 'one Ripped: line, one file'
    Assert-Equal 5.0 $one.RipGB 'ripped GB summed from the Ripped: lines'

    $failedLog = (Get-ChildItem $logDir -Filter 'Slow Film_*' | Sort-Object Name | Select-Object -First 1).FullName
    Assert-True ($null -eq (Get-RipLogSummary -Path $failedLog)) 'a session without RIP SESSION COMPLETE is not summarised'

    $all = @(Get-RipLogSummaries -LogDir $logDir)
    Assert-Equal 7 $all.Count 'Get-RipLogSummaries returns only the 7 completed sessions (no $null entries)'
    Assert-True (-not ($all | Where-Object { $null -eq $_ })) 'no $null in the summaries array'
    $fails = @(Get-RipLogFailures -LogDir $logDir)
    Assert-Equal 1 $fails.Count 'one failed session'
    Assert-Equal 'STEP 1/4: MakeMKV rip' $fails[0].FailedAt 'the failed step is read from Failed at:'

    # ---------------------------------------------------------------------------
    Write-Host "`nEstimated finish" -ForegroundColor Cyan
    $now = [datetime]'2026-10-29 18:11:00'
    $est = Get-RipTimeEstimate -LogDir $logDir -Type 'Movie' -Now $now
    Assert-Equal 6 $est.Count 'six completed Movie rips counted'
    Assert-Equal 52 $est.MedianMinutes 'median of 40,45,50,55,60,235 = 52.5 -> 52 (rounded to even)'
    Assert-Equal '19:03' $est.Finish.ToString('HH:mm') 'finish = now + median'
    Assert-Equal 'Estimated finish: ~19:03 (median of 6 past Movie rips: 52 min, most take 46-59 min)' (Format-RipTimeEstimate $est) 'log line wording'
    Assert-True ($null -eq (Get-RipTimeEstimate -LogDir $logDir -Type 'Documentary' -Now $now)) 'fewer than 3 past rips of the type: no estimate'
    Assert-True ($null -eq (Get-RipTimeEstimate -LogDir $logDir -Type 'Blu-ray' -Now $now)) 'no past rips of the type: no estimate'
    Assert-True ($null -eq (Get-RipTimeEstimate -LogDir (Join-Path $tempRoot 'missing') -Type 'Movie' -Now $now)) 'missing log folder: no estimate, no error'

    # ---------------------------------------------------------------------------
    Write-Host "`nOutliers" -ForegroundColor Cyan
    $out = @(Find-RipOutliers -Sessions $all -Failures $fails)
    Assert-Equal 1 $out.Count 'only the 235-minute film is above the fence'
    Assert-Equal 'Slow Film' $out[0].Title 'the slow film is the outlier'
    Assert-True ([bool]($out[0].Reasons | Where-Object { $_ -like 'waiting: 180 min*' })) 'idle time after encoding is reported as waiting'
    Assert-True ([bool]($out[0].Reasons | Where-Object { $_ -like 'damaged disc: MakeMKV stopped 1x*' })) 'a MakeMKV read-error kill is reported as damage'
    Assert-True ([bool]($out[0].Reasons | Where-Object { $_ -like '1 earlier failed attempt*' })) 'the earlier failed attempt at the same title is linked'
    Assert-True (-not ($out | Where-Object { $_.Type -eq 'Documentary' })) 'a type with under 4 rips is never judged'

    Assert-Equal 52.5 (Get-RipMedian @(40, 45, 50, 55, 60, 235)) 'median of an even count averages the middle two'
    Assert-Equal 46.25 (Get-RipPercentile @(40, 45, 50, 55, 60, 235) 25) 'Q1 interpolated like Excel PERCENTILE.INC'
    Assert-Equal '1h 05m' (Format-RipLogDuration 65) 'durations over an hour read as h/m'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n$script:Passed passed, $script:Failed failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
