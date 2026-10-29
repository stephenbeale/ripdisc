#Requires -Version 5.1
<#
.SYNOPSIS
    How long rips take: per-type averages and outliers from ripdisc and ripaudio logs.

.DESCRIPTION
    Reads every completed session log (RIP SESSION COMPLETE) in the ripdisc log folder
    (C:\Video\logs) and the ripaudio log folder (C:\Music\logs). Prints the median,
    average and range per content type (Movie, Documentary, TV Series, Album ...), the
    MakeMKV / HandBrake / cyanrip phase times, then every unusually slow rip (above
    Q3 + 1.5 x IQR for its type) with the reasons the logs give: waiting at a prompt,
    many titles, a big or damaged disc, slow reads, another rip running at the same
    time, earlier failed attempts. For albums it also reads cyanrip's own log beside the
    album (paranoia re-reads, ripping errors, audio length).

    Read-only: nothing is written except the optional -JsonPath file.

.PARAMETER VideoLogDir
    ripdisc logs. Default: <tempRoot>\logs from ripdisc-config.json, else C:\Video\logs.

.PARAMETER MusicLogDir
    ripaudio logs. Default C:\Music\logs. Pass "" to skip audio.

.PARAMETER JsonPath
    Also write every session and outlier as JSON (for charts).

.EXAMPLE
    .\rip-stats.ps1

.EXAMPLE
    .\rip-stats.ps1 -JsonPath "$env:TEMP\rip-stats.json"
#>
param(
    [string]$VideoLogDir = "",
    [string]$MusicLogDir = "C:\Music\logs",
    [string]$JsonPath = ""
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot "RipStats.ps1")

if (-not $VideoLogDir) {
    $VideoLogDir = "C:\Video\logs"
    $configPath = Join-Path $PSScriptRoot "ripdisc-config.json"
    if (Test-Path -LiteralPath $configPath) {
        try {
            $cfg = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
            if ($cfg.tempRoot) { $VideoLogDir = Join-Path $cfg.tempRoot "logs" }
        } catch { }
    }
}

function Format-Minutes {
    param($Minutes)
    if ($null -eq $Minutes) { return "-" }
    $m = [math]::Round([double]$Minutes)
    if ($m -ge 60) { return "{0}h {1:D2}m" -f [math]::Floor($m / 60), [int]($m % 60) }
    return "$m min"
}

function Get-TypeStats {
    param([object[]]$Sessions)
    foreach ($g in ($Sessions | Group-Object Type | Sort-Object Count -Descending)) {
        $tot = @($g.Group | ForEach-Object { $_.TotalMinutes })
        $rip = @($g.Group | Where-Object { $null -ne $_.RipMinutes } | ForEach-Object { $_.RipMinutes })
        $enc = @($g.Group | Where-Object { $null -ne $_.EncodeMinutes } | ForEach-Object { $_.EncodeMinutes })
        [pscustomobject]@{
            Type = $g.Name; Count = $g.Count
            Median = [math]::Round((Get-RipMedian $tot), 1)
            Average = [math]::Round(($tot | Measure-Object -Average).Average, 1)
            Q1 = [math]::Round((Get-RipPercentile $tot 25), 1); Q3 = [math]::Round((Get-RipPercentile $tot 75), 1)
            Min = [math]::Round(($tot | Measure-Object -Minimum).Minimum, 1)
            Max = [math]::Round(($tot | Measure-Object -Maximum).Maximum, 1)
            MedianRip = if ($rip.Count) { [math]::Round((Get-RipMedian $rip), 1) } else { $null }
            MedianEncode = if ($enc.Count) { [math]::Round((Get-RipMedian $enc), 1) } else { $null }
        }
    }
}

function Show-Section {
    param([string]$Heading, [object[]]$Sessions, [object[]]$Failures, [string]$RipLabel, [switch]$NoEncode)
    Write-Host "`n$Heading" -ForegroundColor Cyan
    if (-not $Sessions -or $Sessions.Count -eq 0) { Write-Host "  no completed sessions" -ForegroundColor Gray; return @() }
    $stats = @(Get-TypeStats $Sessions)
    $cols = @(
        @{ n = 'Type'; e = { $_.Type } }, @{ n = 'Rips'; e = { $_.Count } },
        @{ n = 'Median'; e = { Format-Minutes $_.Median } }, @{ n = 'Average'; e = { Format-Minutes $_.Average } },
        @{ n = 'Most (Q1-Q3)'; e = { "$(Format-Minutes $_.Q1) - $(Format-Minutes $_.Q3)" } },
        @{ n = 'Range'; e = { "$(Format-Minutes $_.Min) - $(Format-Minutes $_.Max)" } },
        @{ n = $RipLabel; e = { Format-Minutes $_.MedianRip } })
    if (-not $NoEncode) { $cols += @{ n = 'HandBrake'; e = { Format-Minutes $_.MedianEncode } } }
    $stats | Format-Table $cols -AutoSize | Out-String -Width 200 | Write-Host

    $outliers = @(Find-RipOutliers -Sessions $Sessions -Failures $Failures)
    Write-Host "  Unusually slow (above Q3 + 1.5 x IQR for the type): $($outliers.Count)" -ForegroundColor Yellow
    foreach ($o in $outliers) {
        Write-Host ("  [{0}] {1} - {2} (typical {3}) {4:yyyy-MM-dd}" -f $o.Type, $o.Title, (Format-Minutes $o.TotalMinutes), (Format-Minutes $o.MedianMinutes), $o.Start) -ForegroundColor White
        foreach ($r in $o.Reasons) { Write-Host "      - $r" -ForegroundColor Gray }
    }
    return [pscustomobject]@{ Stats = $stats; Outliers = $outliers }
}

Write-Host "Reading $VideoLogDir ..." -ForegroundColor Gray
$video = @(Get-RipLogSummaries -LogDir $VideoLogDir)
$videoFailures = @(Get-RipLogFailures -LogDir $VideoLogDir)
$audio = @(); $audioFailures = @()
if ($MusicLogDir) {
    Write-Host "Reading $MusicLogDir (and each album's cyanrip log) ..." -ForegroundColor Gray
    $audio = @(Get-RipLogSummaries -LogDir $MusicLogDir -IncludeCyanrip)
    $audioFailures = @(Get-RipLogFailures -LogDir $MusicLogDir)
}

$videoReport = Show-Section -Heading "ripdisc - $($video.Count) completed rips ($($videoFailures.Count) failed/aborted sessions not counted)" `
    -Sessions $video -Failures $videoFailures -RipLabel 'MakeMKV'
$audioReport = $null
if ($MusicLogDir) {
    $audioReport = Show-Section -Heading "ripaudio - $($audio.Count) completed albums ($($audioFailures.Count) failed/aborted sessions not counted)" `
        -Sessions $audio -Failures $audioFailures -RipLabel 'cyanrip' -NoEncode
}

Write-Host "`nTimes are session start to RIP SESSION COMPLETE, so they include time spent waiting at prompts." -ForegroundColor Gray
Write-Host "The median is the better guide; a few long waits pull the average up." -ForegroundColor Gray

if ($JsonPath) {
    $export = [pscustomobject]@{
        Generated = (Get-Date).ToString('s')
        VideoLogDir = $VideoLogDir; MusicLogDir = $MusicLogDir
        Video = [pscustomobject]@{ Stats = $videoReport.Stats; Outliers = $videoReport.Outliers; FailedSessions = $videoFailures.Count
            Sessions = @($video | Select-Object Type, Title, Log, Start, TotalMinutes, SetupMinutes, RipMinutes, EncodeMinutes, Files, RipGB, OutputGB, ReadErrorKills, DecoderErrors) }
        Audio = [pscustomobject]@{ Stats = $audioReport.Stats; Outliers = $audioReport.Outliers; FailedSessions = $audioFailures.Count
            Sessions = @($audio | Select-Object Type, Title, Log, Start, TotalMinutes, SetupMinutes, RipMinutes, PostRipMinutes, Tracks, SizeMB, AccurateRip, DataErrorTracks,
                @{ n = 'AudioMinutes'; e = { $_.Cyanrip.AudioMinutes } }, @{ n = 'RereadPct'; e = { if ($_.Cyanrip.Read) { [math]::Round(100 * $_.Cyanrip.Verify / $_.Cyanrip.Read, 1) } } },
                @{ n = 'Fixups'; e = { $_.Cyanrip.Fixups } }, @{ n = 'CyanripErrors'; e = { $_.Cyanrip.Errors } }) }
    }
    $export | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $JsonPath -Encoding UTF8
    Write-Host "JSON written: $JsonPath" -ForegroundColor Green
}
