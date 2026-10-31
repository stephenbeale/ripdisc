#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the auto-discovery disc scan: MakeMKV robot-line parsing and the scan runner
    with its live status line and timeout.

.DESCRIPTION
    Function bodies are lifted out of rip-disc.ps1 with the AST parser, so the tests cannot
    drift from the shipped code. Invoke-MakeMkvInfoWithStatus runs against a fake makemkvcon
    (a .cmd file that prints captured-style robot output), started as a normal child process
    (no -EncodedCommand: AVG deletes scripts that do that, see CLAUDE.md).

.EXAMPLE
    .\tests\Test-DiscoveryScan.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ScriptPath has parse errors" }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

foreach ($f in 'ConvertFrom-MakeMkvInfoLine', 'Invoke-MakeMkvInfoWithStatus', 'Format-ScanElapsed', 'Get-DiscInfo') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $f)))
}

$script:Passed = 0; $script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because`n        expected: [$Expected]`n        actual:   [$Actual]" -ForegroundColor Red }
}
function Assert-True { param([bool]$Condition, [string]$Because) Assert-Equal $true $Condition $Because }

Write-Host "`nConvertFrom-MakeMkvInfoLine" -ForegroundColor Cyan
$op = ConvertFrom-MakeMkvInfoLine 'PRGC:5018,0,"Opening DVD disc"'
Assert-Equal 'Operation' $op.Kind "PRGC is an operation"
Assert-Equal 'Opening DVD disc' $op.Text "operation name extracted"
Assert-Equal 'Operation' (ConvertFrom-MakeMkvInfoLine 'PRGT:5018,0,"Scanning CD-ROM devices"').Kind "PRGT is an operation"
$pv = ConvertFrom-MakeMkvInfoLine 'PRGV:100,32768,65536'
Assert-Equal 'Progress' $pv.Kind "PRGV is progress"
Assert-Equal 50 $pv.Percent "percent is total/max"
Assert-Equal $null (ConvertFrom-MakeMkvInfoLine 'PRGV:0,0,0') "max 0 is ignored, no divide by zero"
Assert-Equal 'TitleAdded' (ConvertFrom-MakeMkvInfoLine 'MSG:3307,0,2,"Title #1 was added (12 cell(s), 2:29:22)","Title #%1 was added (%2 cell(s), %3)","1","12","2:29:22"').Kind "title-added message is counted"
Assert-Equal 'Message' (ConvertFrom-MakeMkvInfoLine 'MSG:1005,0,1,"MakeMKV v1.18.4 win(x64-release) started","%1 started","x"').Kind "other MSG is a message"
Assert-Equal $null (ConvertFrom-MakeMkvInfoLine 'TINFO:0,9,0,"2:29:22"') "TINFO carries no status"
Assert-Equal $null (ConvertFrom-MakeMkvInfoLine '') "blank line"

Write-Host "`nFormat-ScanElapsed" -ForegroundColor Cyan
Assert-Equal '0:07' (Format-ScanElapsed ([TimeSpan]::FromSeconds(7))) "seconds padded"
Assert-Equal '3:05' (Format-ScanElapsed ([TimeSpan]::FromSeconds(185))) "minutes and seconds"
Assert-Equal '75:00' (Format-ScanElapsed ([TimeSpan]::FromMinutes(75))) "over an hour stays in minutes"

# ---- fake makemkvcon ----
$tmp = Join-Path ([IO.Path]::GetTempPath()) "ripdisc-discovery-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    $fake = Join-Path $tmp 'fake-makemkvcon.cmd'
    @'
@echo off
echo MSG:1005,0,1,"MakeMKV v1.18.4 win(x64-release) started","%%1 started","x"
echo PRGT:5018,0,"Opening DVD disc"
echo PRGV:0,0,65536
ping -n 2 127.0.0.1 >nul
echo PRGC:5018,0,"Processing title sets"
echo MSG:3307,0,2,"Title #1 was added (12 cell(s), 2:29:22)","x","1","12","2:29:22"
echo PRGV:0,32768,65536
echo MSG:3307,0,2,"Title #2 was added (1 cell(s), 0:03:10)","x","2","1","0:03:10"
echo CINFO:1,6206,"DVD disc"
echo CINFO:2,0,"THE_DIRTY_DOZEN"
echo CINFO:32,0,"THE_DIRTY_DOZEN"
echo TCOUNT:2
echo TINFO:0,9,0,"2:29:22"
echo TINFO:1,9,0,"0:03:10"
exit /b 0
'@ | Set-Content -Path $fake -Encoding ASCII

    $hang = Join-Path $tmp 'hang-makemkvcon.cmd'
    @'
@echo off
echo PRGT:5018,0,"Opening DVD disc"
ping -n 8 127.0.0.1 >nul
exit /b 0
'@ | Set-Content -Path $hang -Encoding ASCII

    Write-Host "`nInvoke-MakeMkvInfoWithStatus - normal scan" -ForegroundColor Cyan
    $makemkvconPath = $fake
    $r = Invoke-MakeMkvInfoWithStatus -DiscSource 'disc:1' -TimeoutSec 30 6>$null
    Assert-True ($null -ne $r) "returns a result"
    Assert-Equal 0 $r.ExitCode "exit code passed through"
    Assert-True (@($r.Lines | Where-Object { $_ -like 'CINFO:2,*' }).Count -eq 1) "all output lines captured"

    Write-Host "`nGet-DiscInfo on the fake scan" -ForegroundColor Cyan
    $info = Get-DiscInfo -DiscSource 'disc:1' 6>$null
    Assert-Equal 'DVD disc' $info.DiscType "disc type parsed"
    Assert-Equal 'THE_DIRTY_DOZEN' $info.DiscName "disc name parsed"
    Assert-Equal 2 $info.Titles.Count "titles parsed"

    Write-Host "`nInvoke-MakeMkvInfoWithStatus - hung drive hits the cap" -ForegroundColor Cyan
    $makemkvconPath = $hang
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r2 = Invoke-MakeMkvInfoWithStatus -DiscSource 'disc:1' -TimeoutSec 2 6>$null
    $sw.Stop()
    Assert-Equal $null $r2 "timeout returns null"
    Assert-True ($sw.Elapsed.TotalSeconds -lt 6) "gave up near the cap, not when the fake finished ($([int]$sw.Elapsed.TotalSeconds)s)"
} finally {
    Start-Sleep -Milliseconds 300
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}

foreach ($f in 'Read-BigEndianUInt', 'Get-DvdTitleDurations', 'Get-BlurayPlaylistDurations', 'Get-DiscTitleDurations', 'ConvertTo-DurationSeconds', 'Get-SeriesLikelihood') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $f)))
}

Write-Host "`nGet-SeriesLikelihood" -ForegroundColor Cyan
$m = { param($min) [int]($min * 60) }
$h = Get-SeriesLikelihood -DurationsSec @((& $m 143.5), 12, (& $m 9.2), (& $m 3.5))
Assert-True (-not $h.LooksLikeSeries) "The Dirty Dozen (real disc 2026-10-31): a film, not a series"
$h = Get-SeriesLikelihood -DurationsSec @((& $m 22), (& $m 21.6), (& $m 23.1), (& $m 22.4), (& $m 4))
Assert-True $h.LooksLikeSeries "four ~22-min titles: series"
Assert-Equal 4 $h.EpisodeCount "episode count"
Assert-Equal 22 $h.TypicalMinutes "typical length"
$h = Get-SeriesLikelihood -DurationsSec @((& $m 44), (& $m 45), (& $m 43), (& $m 132))
Assert-True $h.LooksLikeSeries "three ~44-min episodes plus a play-all of them: series"
$h = Get-SeriesLikelihood -DurationsSec @((& $m 118), (& $m 24), (& $m 22), (& $m 25))
Assert-True (-not $h.LooksLikeSeries) "film plus three ~24-min featurettes: the feature is no play-all, so not a series"
$h = Get-SeriesLikelihood -DurationsSec @((& $m 54), (& $m 57), (& $m 68))
Assert-True (-not $h.LooksLikeSeries) "lengths too far apart (54/57/68): no hint"
$h = Get-SeriesLikelihood -DurationsSec @((& $m 30), (& $m 30), (& $m 30), (& $m 30), (& $m 150))
Assert-True (-not $h.LooksLikeSeries) "identical copies count once (copy-protection padding)"
Assert-True (-not (Get-SeriesLikelihood -DurationsSec @()).LooksLikeSeries) "no lengths: no hint"

Write-Host "`nConvertTo-DurationSeconds" -ForegroundColor Cyan
Assert-Equal 5362 (ConvertTo-DurationSeconds '1:29:22') "h:mm:ss"
Assert-Equal 190 (ConvertTo-DurationSeconds '3:10') "mm:ss"
Assert-Equal 0 (ConvertTo-DurationSeconds 'junk') "unparseable is 0"

Write-Host "`nIFO / MPLS readers on synthetic files" -ForegroundColor Cyan
$tmp2 = Join-Path ([IO.Path]::GetTempPath()) "ripdisc-ifo-$([guid]::NewGuid().ToString('N'))"
try {
    $vts = Join-Path $tmp2 'VIDEO_TS'; New-Item -ItemType Directory -Path $vts -Force | Out-Null
    # VTS IFO: signature, VTS_PGCI at sector 1 (0xCC), 3 PGCs: entry 0:44:30, non-entry 0:01:00, entry 1:02:03
    $b = New-Object byte[] 4096
    [Text.Encoding]::ASCII.GetBytes('DVDVIDEO-VTS').CopyTo($b, 0)
    $b[0xCF] = 1
    $pg = 2048; $b[$pg + 1] = 3
    $pgcOff = @(0x40, 0x80, 0xC0); $entry = @(0x81, 0x00, 0x82); $times = @(@(0x00, 0x44, 0x30), @(0x00, 0x01, 0x00), @(0x01, 0x02, 0x03))
    for ($i = 0; $i -lt 3; $i++) {
        $e = $pg + 8 + 8 * $i
        $b[$e] = $entry[$i]; $b[$e + 7] = $pgcOff[$i]
        $p = $pg + $pgcOff[$i]; $b[$p + 4] = $times[$i][0]; $b[$p + 5] = $times[$i][1]; $b[$p + 6] = $times[$i][2]; $b[$p + 7] = 0xC0
    }
    [IO.File]::WriteAllBytes((Join-Path $vts 'VTS_01_0.IFO'), $b)
    [IO.File]::WriteAllBytes((Join-Path $vts 'VTS_02_0.IFO'), [byte[]](1..50))   # junk: skipped
    $d = @(Get-DiscTitleDurations -DriveRoot $tmp2)
    Assert-Equal '2670,3723' ($d -join ',') "DVD: entry PGCs only, BCD times, junk IFO skipped"

    $pl = Join-Path $tmp2 'BD\BDMV\PLAYLIST'; New-Item -ItemType Directory -Path $pl -Force | Out-Null
    # MPLS: 2 play items, 600 s + 720 s at 45 kHz
    $mb = New-Object byte[] 200
    [Text.Encoding]::ASCII.GetBytes('MPLS0200').CopyTo($mb, 0)
    $mb[11] = 40; $mb[40 + 7] = 2
    $put = { param($o, [uint32]$v) $mb[$o] = ($v -shr 24) -band 0xFF; $mb[$o + 1] = ($v -shr 16) -band 0xFF; $mb[$o + 2] = ($v -shr 8) -band 0xFF; $mb[$o + 3] = $v -band 0xFF }
    $item = 50
    foreach ($secs in 600, 720) {
        $mb[$item + 1] = 20                       # item length 20 (+2 for the length field)
        & $put ($item + 14) 1000
        & $put ($item + 18) (1000 + $secs * 45000)
        $item += 22
    }
    [IO.File]::WriteAllBytes((Join-Path $pl '00001.mpls'), $mb)
    $d = @(Get-DiscTitleDurations -DriveRoot (Join-Path $tmp2 'BD'))
    Assert-Equal '1320' ($d -join ',') "Blu-ray: playlist length is the sum of its play items"
    Assert-Equal 0 @(Get-DiscTitleDurations -DriveRoot (Join-Path $tmp2 'nothing')).Count "no disc structure: no lengths"
} finally {
    Remove-Item -Recurse -Force $tmp2 -ErrorAction SilentlyContinue
}

Write-Host "`n$($script:Passed) passed, $($script:Failed) failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed -gt 0) { exit 1 }
