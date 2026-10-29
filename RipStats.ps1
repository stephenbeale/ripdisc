# Rip-time history read back from session logs: ripdisc's C:\Video\logs and, for the
# stats report, ripaudio's C:\Music\logs (same "[yyyy-MM-dd HH:mm:ss] message" format).
# Dot-sourced by rip-disc.ps1 (estimated finish line) and rip-stats.ps1 (report).
# Read-only: never writes to a log or a media folder.

function ConvertFrom-RipLogStamp {
    param([string]$Line)
    if ($Line -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]') {
        return [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    }
    return $null
}

# First line matching $Pattern, as its timestamp; $null when absent.
function Get-RipLogStampOf {
    param([string[]]$Lines, [string]$Pattern, [switch]$Last)
    $hits = @($Lines | Where-Object { $_ -match $Pattern })
    if ($hits.Count -eq 0) { return $null }
    $pick = if ($Last) { $hits[-1] } else { $hits[0] }
    return ConvertFrom-RipLogStamp $pick
}

function Get-RipLogValue {
    param([string[]]$Lines, [string]$Label)
    $hit = $Lines | Where-Object { $_ -match "^\[[^\]]+\] $([regex]::Escape($Label)): (.*)$" } | Select-Object -First 1
    if ($hit -and $hit -match "^\[[^\]]+\] $([regex]::Escape($Label)): (.*)$") { return $Matches[1].Trim() }
    return $null
}

function Get-RipMinutes {
    param($From, $To)
    if ($null -eq $From -or $null -eq $To) { return $null }
    return [math]::Round(($To - $From).TotalMinutes, 1)
}

# cyanrip writes its own log next to the album ("<Album>.log", first line "cyanrip x.y").
# Albums are sometimes ripped to C: and copied to F:, so the other drive is tried too.
function Get-CyanripLogInfo {
    param([string]$AlbumDir)
    if (-not $AlbumDir) { return $null }
    $dirs = @($AlbumDir)
    if ($AlbumDir -match '^[A-Za-z]:(.*)$') { $dirs += @("F:$($Matches[1])", "C:$($Matches[1])") }
    foreach ($dir in ($dirs | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
        $log = Get-ChildItem -LiteralPath $dir -Filter *.log -File -ErrorAction SilentlyContinue |
            Where-Object { (Get-Content -LiteralPath $_.FullName -TotalCount 1) -like 'cyanrip *' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $log) { continue }
        $c = Get-Content -LiteralPath $log.FullName
        $info = [ordered]@{ Path = $log.FullName; AudioMinutes = $null; Read = 0; Verify = 0; Fixups = 0;
            Skips = 0; Errors = $null; FinishedAt = $null; AccurateRip = $null; Paranoia = $null; DriveModel = $null }
        foreach ($line in $c) {
            if ($null -eq $info.AudioMinutes -and $line -match '^Total time:\s+(\d+):(\d+):([\d.]+)') {
                $info.AudioMinutes = [math]::Round([int]$Matches[1] * 60 + [int]$Matches[2] + [double]$Matches[3] / 60, 1)
            }
            elseif (-not $info.Paranoia -and $line -match '^Paranoia level:\s+(.+)$') { $info.Paranoia = $Matches[1].Trim() }
            elseif (-not $info.DriveModel -and $line -match '^Device model:\s+(.+)$') { $info.DriveModel = ($Matches[1] -replace '\s+', ' ').Trim() }
            elseif (-not $info.AccurateRip -and $line -match '^AccurateRip:\s+(.+)$') { $info.AccurateRip = $Matches[1].Trim() }
            elseif ($line -match '^Ripping errors:\s+(\d+)') { $info.Errors = [int]$Matches[1] }
            elseif ($line -match '^Ripping finished at (\S+)') {
                $t = [datetime]::MinValue
                if ([datetime]::TryParse($Matches[1], [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$t)) { $info.FinishedAt = $t }
            }
        }
        # The album-level "Paranoia status counts" block is the last one in the file.
        $start = ($c | Select-String -SimpleMatch 'Paranoia status counts:' | Select-Object -Last 1).LineNumber
        if ($start) {
            foreach ($line in $c[$start..([math]::Min($start + 12, $c.Count - 1))]) {
                if ($line -match '^\s+([A-Z_]+):\s+(\d+)') {
                    # switch -Regex resets $Matches, so take the values first
                    $name = $Matches[1]; $count = [int]$Matches[2]
                    switch -Regex ($name) {
                        '^READ$'   { $info.Read = $count }
                        '^VERIFY$' { $info.Verify = $count }
                        '^FIXUP'   { $info.Fixups += $count }
                        '^(SKIP|SCRATCH|DRIFT)$' { $info.Skips += $count }
                    }
                }
            }
        }
        return [pscustomobject]$info
    }
    return $null
}

# One completed session log as an object, or $null for a session that never reached
# "RIP SESSION COMPLETE" (failed, aborted or still running).
function Get-RipLogSummary {
    param([Parameter(Mandatory)][string]$Path, [switch]$IncludeCyanrip)
    $c = @(Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue)
    if ($c.Count -eq 0 -or -not ($c -match 'RIP SESSION COMPLETE')) { return $null }

    $start = ConvertFrom-RipLogStamp $c[0]
    $end = Get-RipLogStampOf $c 'RIP SESSION COMPLETE'
    $ripStart = Get-RipLogStampOf $c 'STEP 1/4: Starting'
    $ripEnd = Get-RipLogStampOf $c 'STEP 1/4: .*complete'
    $isAudio = [bool](Get-RipLogValue $c 'Album')

    $s = [ordered]@{
        Kind = if ($isAudio) { 'Audio' } else { 'Video' }
        Type = $null; Title = $null; Log = Split-Path $Path -Leaf; Drive = Get-RipLogValue $c 'Drive'
        Start = $start; End = $end
        TotalMinutes = Get-RipMinutes $start $end
        SetupMinutes = Get-RipMinutes $start $ripStart
        RipMinutes = Get-RipMinutes $ripStart $ripEnd
        EncodeMinutes = $null; Files = 0; RipGB = 0.0; OutputGB = $null
        ReadErrorKills = @($c -match 'MakeMKV killed after \d+ consecutive read errors').Count
        DecoderErrors = 0; EjectRetries = @($c -match 'Retrying disc eject').Count
        DataErrorTracks = $null; AccurateRip = $null; Tracks = $null; SizeMB = $null
        PostRipMinutes = $null; Cyanrip = $null
    }

    if ($isAudio) {
        $artist = Get-RipLogValue $c 'Artist'
        $album = Get-RipLogValue $c 'Album'
        $s.Type = 'Album'
        $s.Title = if ($artist) { "$artist - $album" } else { $album }
        $tracks = $c | Where-Object { $_ -match 'Total tracks: (\d+)' } | Select-Object -Last 1
        if ($tracks -match 'Total tracks: (\d+)') { $s.Tracks = [int]$Matches[1] }
        $size = $c | Where-Object { $_ -match 'Total size: ([\d.]+) MB' } | Select-Object -Last 1
        if ($size -match 'Total size: ([\d.]+) MB') { $s.SizeMB = [double]$Matches[1] }
        $s.DataErrorTracks = Get-RipLogValue $c 'Data errors detected on tracks'
        $ar = $c | Where-Object { $_ -match 'AccurateRip: (\d+/\d+)' } | Select-Object -Last 1
        if ($ar -match 'AccurateRip: (\d+/\d+)') { $s.AccurateRip = $Matches[1] }
        elseif ($v = Get-RipLogValue $c 'AccurateRip DB status') { $s.AccurateRip = "DB $v" }
        $s.PostRipMinutes = Get-RipMinutes $ripEnd $end
        $s.Files = $s.Tracks
        if ($IncludeCyanrip) { $s.Cyanrip = Get-CyanripLogInfo (Get-RipLogValue $c 'Final Output') }
    } else {
        $s.Type = Get-RipLogValue $c 'Type'
        $s.Title = Get-RipLogValue $c 'Title'
        $encStart = Get-RipLogStampOf $c 'STEP 2/4: Starting'
        $encEnd = Get-RipLogStampOf $c 'STEP 2/4: .*complete'
        $s.EncodeMinutes = Get-RipMinutes $encStart $encEnd
        # "  Ripped: B1_t00.mkv (5.67 GB)" lines are Step 1's MakeMKV output files.
        foreach ($line in ($c -match '^\[[^\]]+\]\s+Ripped: .+\(([\d.]+) GB\)$')) {
            if ($line -match '\(([\d.]+) GB\)$') { $s.Files++; $s.RipGB += [double]$Matches[1] }
        }
        $s.RipGB = [math]::Round($s.RipGB, 2)
        $out = $c | Where-Object { $_ -match 'Total size: ([\d.]+) GB' } | Select-Object -Last 1
        if ($out -match 'Total size: ([\d.]+) GB') { $s.OutputGB = [double]$Matches[1] }
        foreach ($line in ($c -match 'decoder done: \d+ frames, (\d+) decoder errors')) {
            if ($line -match ', (\d+) decoder errors') { $s.DecoderErrors += [int]$Matches[1] }
        }
    }
    return [pscustomobject]$s
}

# Every completed session in $LogDir. Skips ripaudio's metadata-tool and continue logs,
# which are not rip sessions.
function Get-RipLogSummaries {
    param([Parameter(Mandatory)][string]$LogDir, [switch]$IncludeCyanrip)
    if (-not (Test-Path -LiteralPath $LogDir -PathType Container)) { return @() }
    $logs = Get-ChildItem -LiteralPath $LogDir -Filter *.log -File |
        Where-Object { $_.Name -notmatch '^(audit|get|search|undo)-metadata_' -and $_.Name -notmatch '_continue_' }
    return @($logs | ForEach-Object { Get-RipLogSummary -Path $_.FullName -IncludeCyanrip:$IncludeCyanrip } | Where-Object { $null -ne $_ })
}

# Failed/aborted sessions (no RIP SESSION COMPLETE) as Title + Start, so a slow success
# can be linked to earlier failed attempts at the same disc.
function Get-RipLogFailures {
    param([Parameter(Mandatory)][string]$LogDir)
    if (-not (Test-Path -LiteralPath $LogDir -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $LogDir -Filter *.log -File | ForEach-Object {
        $c = @(Get-Content -LiteralPath $_.FullName -ErrorAction SilentlyContinue)
        if ($c.Count -gt 0 -and -not ($c -match 'RIP SESSION COMPLETE')) {
            $artist = Get-RipLogValue $c 'Artist'; $album = Get-RipLogValue $c 'Album'
            $title = if ($album) { if ($artist) { "$artist - $album" } else { $album } } else { Get-RipLogValue $c 'Title' }
            $failed = ($c -match 'Failed at: ' | Select-Object -First 1) -replace '^\[[^\]]+\] Failed at: ', ''
            [pscustomobject]@{ Title = $title; Start = ConvertFrom-RipLogStamp $c[0]; FailedAt = $failed
                LastStamp = ConvertFrom-RipLogStamp (@($c | Where-Object { $_ -match '^\[' })[-1]) }
        }
    })
}

function Format-RipLogDuration {
    param([double]$Minutes)
    $m = [math]::Round($Minutes)
    if ($m -ge 60) { return "{0}h {1:D2}m" -f [math]::Floor($m / 60), [int]($m % 60) }
    return "$m min"
}

function Get-RipMedian {
    param([double[]]$Values)
    $v = @($Values | Sort-Object)
    if ($v.Count -eq 0) { return $null }
    $mid = [math]::Floor($v.Count / 2)
    if ($v.Count % 2) { return $v[$mid] }
    return ($v[$mid - 1] + $v[$mid]) / 2
}

# Linear-interpolated percentile (0-100), as Excel's PERCENTILE.INC.
function Get-RipPercentile {
    param([double[]]$Values, [double]$P)
    $v = @($Values | Sort-Object)
    if ($v.Count -eq 0) { return $null }
    $rank = ($P / 100) * ($v.Count - 1)
    $lo = [math]::Floor($rank); $hi = [math]::Ceiling($rank)
    return $v[$lo] + ($v[$hi] - $v[$lo]) * ($rank - $lo)
}

# Median past session length for this content type, for the "Estimated finish" log line.
# $null with fewer than $MinSamples completed sessions of that type.
function Get-RipTimeEstimate {
    param([Parameter(Mandatory)][string]$LogDir, [Parameter(Mandatory)][string]$Type,
        [datetime]$Now = (Get-Date), [int]$MinSamples = 3)
    if (-not (Test-Path -LiteralPath $LogDir -PathType Container)) { return $null }
    # Only parse logs of this type: rip-disc.ps1 calls this at the start of every rip
    $paths = @(Select-String -Path (Join-Path $LogDir '*.log') -Pattern "^\[[^\]]+\] Type: $([regex]::Escape($Type))$" -List |
        ForEach-Object { $_.Path })
    $same = @($paths | ForEach-Object { Get-RipLogSummary -Path $_ } |
        Where-Object { $null -ne $_ -and $_.Kind -eq 'Video' -and $_.Type -eq $Type -and $_.TotalMinutes -gt 0 })
    if ($same.Count -lt $MinSamples) { return $null }
    $median = Get-RipMedian ($same | ForEach-Object { $_.TotalMinutes })
    return [pscustomobject]@{
        Type = $Type; Count = $same.Count; MedianMinutes = [math]::Round($median)
        Low = [math]::Round((Get-RipPercentile ($same | ForEach-Object { $_.TotalMinutes }) 25))
        High = [math]::Round((Get-RipPercentile ($same | ForEach-Object { $_.TotalMinutes }) 75))
        Finish = $Now.AddMinutes($median)
    }
}

function Format-RipTimeEstimate {
    param([Parameter(Mandatory)]$Estimate)
    return "Estimated finish: ~$($Estimate.Finish.ToString('HH:mm')) (median of $($Estimate.Count) past $($Estimate.Type) rips: $($Estimate.MedianMinutes) min, most take $($Estimate.Low)-$($Estimate.High) min)"
}

# Sessions whose time span overlaps $Session's (other drives ripping/encoding at once).
function Get-RipOverlaps {
    param($Session, [object[]]$All)
    @($All | Where-Object { $_.Log -ne $Session.Log -and $_.Start -lt $Session.End -and $_.End -gt $Session.Start })
}

# Upper-fence outliers (Q3 + 1.5 x IQR on TotalMinutes) per Type, with plain-English
# reasons drawn from what the logs record. Groups with fewer than 4 sessions are skipped.
function Find-RipOutliers {
    param([Parameter(Mandatory)][object[]]$Sessions, [object[]]$Failures = @())
    $everything = $Sessions
    foreach ($g in ($Sessions | Group-Object Type)) {
        if ($g.Count -lt 4) { continue }
        $tot = @($g.Group | ForEach-Object { $_.TotalMinutes })
        $q1 = Get-RipPercentile $tot 25; $q3 = Get-RipPercentile $tot 75
        $fence = $q3 + 1.5 * ($q3 - $q1)
        $medRip = Get-RipMedian @($g.Group | Where-Object { $null -ne $_.RipMinutes } | ForEach-Object { $_.RipMinutes })
        $medEnc = Get-RipMedian @($g.Group | Where-Object { $null -ne $_.EncodeMinutes } | ForEach-Object { $_.EncodeMinutes })
        $medFiles = Get-RipMedian @($g.Group | ForEach-Object { [double]$_.Files })
        $medGB = Get-RipMedian @($g.Group | Where-Object { $_.RipGB -gt 0 } | ForEach-Object { $_.RipGB })
        $medRipPerGB = Get-RipMedian @($g.Group | Where-Object { $_.RipGB -gt 0.5 -and $_.RipMinutes } | ForEach-Object { $_.RipMinutes / $_.RipGB })
        $medEncPerGB = Get-RipMedian @($g.Group | Where-Object { $_.RipGB -gt 0.5 -and $_.EncodeMinutes } | ForEach-Object { $_.EncodeMinutes / $_.RipGB })
        $medRipPerAudio = Get-RipMedian @($g.Group | Where-Object { $_.Cyanrip.AudioMinutes -gt 0 -and $_.RipMinutes } | ForEach-Object { $_.RipMinutes / $_.Cyanrip.AudioMinutes })

        foreach ($s in ($g.Group | Where-Object { $_.TotalMinutes -gt $fence } | Sort-Object TotalMinutes -Descending)) {
            $why = [System.Collections.Generic.List[string]]::new()
            $work = 0 + $s.RipMinutes + $s.EncodeMinutes
            if ($s.Kind -eq 'Video') {
                $idle = $s.TotalMinutes - $work
                if ($idle -gt 20 -and $idle -gt 0.3 * $s.TotalMinutes) { $why.Add("waiting: $([math]::Round($idle)) min outside ripping and encoding (a prompt or a later step sat waiting)") }
                if ($medFiles -and $s.Files -ge [math]::Max(4, 2 * $medFiles)) { $why.Add("lots of titles: $($s.Files) files ripped (group median $medFiles) - extras or episodes") }
                if ($medGB -and $s.RipGB -ge 1.5 * $medGB) { $why.Add("big disc: $($s.RipGB) GB ripped (median $([math]::Round($medGB,1)) GB) - long running time or high bitrate") }
                if ($s.RipGB -gt 0.5 -and $medRipPerGB -and $s.RipMinutes / $s.RipGB -gt 2 * $medRipPerGB) { $why.Add("slow read: $([math]::Round($s.RipMinutes / $s.RipGB,1)) min/GB vs median $([math]::Round($medRipPerGB,1)) - scratched disc or struggling drive") }
                if ($s.RipGB -gt 0.5 -and $medEncPerGB -and $s.EncodeMinutes / $s.RipGB -gt 2 * $medEncPerGB) { $why.Add("slow encode: $([math]::Round($s.EncodeMinutes / $s.RipGB,1)) min/GB vs median $([math]::Round($medEncPerGB,1))") }
                elseif ($medEnc -and $s.EncodeMinutes -gt 2.5 * $medEnc -and $why.Count -eq 0) { $why.Add("long encode: $([math]::Round($s.EncodeMinutes)) min vs median $([math]::Round($medEnc))") }
                if ($s.ReadErrorKills) { $why.Add("damaged disc: MakeMKV stopped $($s.ReadErrorKills)x on consecutive read errors") }
                if ($s.DecoderErrors -gt 0) { $why.Add("damaged disc: $($s.DecoderErrors) HandBrake decoder errors") }
                if ($s.EjectRetries) { $why.Add("eject needed $($s.EjectRetries) retries") }
                if ($s.Type -eq 'Blu-ray') { $why.Add('Blu-ray: about 5x the data of a DVD, so encoding takes much longer') }
            } else {
                $c = $s.Cyanrip
                if ($s.PostRipMinutes -gt 10) { $why.Add("waiting: $([math]::Round($s.PostRipMinutes)) min after cyanrip finished (a prompt, e.g. the data-error choice)") }
                # Minutes between cyanrip's own "Ripping finished at" and the script logging
                # Step 1 complete: the window sat waiting, not ripping.
                $tail = 0
                if ($c -and $c.FinishedAt -and $s.RipMinutes) {
                    $tail = [math]::Max(0, ($s.Start.AddMinutes($s.SetupMinutes + $s.RipMinutes) - $c.FinishedAt).TotalMinutes)
                    if ($tail -gt 15) { $why.Add("idle: cyanrip's own log says it finished at $($c.FinishedAt.ToString('HH:mm')), but the script only moved on $(Format-RipLogDuration $tail) later (sat waiting, not ripping)") }
                }
                if ($s.SetupMinutes -gt 10) { $why.Add("waiting: $([math]::Round($s.SetupMinutes)) min before ripping started (release picker / prompts)") }
                if ($c -and $c.Read -gt 0) {
                    $re = [math]::Round(100 * $c.Verify / $c.Read)
                    # ~20% VERIFY is cyanrip's normal baseline at max paranoia (172 albums: 17-22%)
                    if ($re -ge 25) { $why.Add("re-reads: paranoia had to verify $re% of sectors (normal is about 20%) - scratched or dirty disc") }
                    # Most clean discs log 0-50 fix-ups; damaged ones run to thousands or millions
                    if ($c.Fixups -gt 1000) { $why.Add("damaged disc: $($c.Fixups.ToString('N0')) paranoia fix-ups (a clean disc has under 50)") }
                    if ($c.Skips -gt 0) { $why.Add("$($c.Skips) unreadable sectors skipped") }
                }
                if ($c -and $c.Errors -gt 0) { $why.Add("cyanrip reported $($c.Errors) ripping errors") }
                if ($s.DataErrorTracks) { $why.Add("data errors on track(s) $($s.DataErrorTracks)") }
                if ($s.AccurateRip -match '^0/') { $why.Add("AccurateRip verified $($s.AccurateRip) tracks (in the database but no matches)") }
                if ($c -and $c.AudioMinutes -gt 72) { $why.Add("long disc: $($c.AudioMinutes) min of audio") }
                if ($c -and $c.AudioMinutes -gt 0 -and $medRipPerAudio -and ($s.RipMinutes - $tail) / $c.AudioMinutes -gt 2 * $medRipPerAudio) {
                    $why.Add("slow read: $(Format-RipLogDuration ($s.RipMinutes - $tail)) to read $([math]::Round($c.AudioMinutes)) min of audio ($([math]::Round(($s.RipMinutes - $tail) / $c.AudioMinutes,1))x real time; typical $([math]::Round($medRipPerAudio,1))x)$(if ($c.DriveModel) { " on the $($c.DriveModel)" })")
                }
            }
            $others = @(Get-RipOverlaps $s $everything)
            if ($others.Count) { $why.Add("ran alongside $($others.Count) other rip(s): $((($others | Select-Object -First 3) | ForEach-Object { $_.Title }) -join '; ')") }
            $earlier = @($Failures | Where-Object { $_.Title -eq $s.Title -and $_.Start -lt $s.Start })
            if ($earlier.Count) { $why.Add("$($earlier.Count) earlier failed attempt(s) at this disc$(if ($earlier[-1].FailedAt) { " (last: $($earlier[-1].FailedAt))" })") }
            if ($why.Count -eq 0) { $why.Add('nothing in the logs explains it') }

            [pscustomobject]@{ Type = $g.Name; Title = $s.Title; Log = $s.Log; Start = $s.Start
                TotalMinutes = $s.TotalMinutes; FenceMinutes = [math]::Round($fence)
                MedianMinutes = [math]::Round((Get-RipMedian $tot)); Reasons = @($why) }
        }
    }
}
