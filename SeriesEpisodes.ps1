# SeriesEpisodes.ps1 - plain -Series episode numbering, extras classification,
# rename manifest and optional TMDb season lookup.
#
# Dot-sourced by both rip-disc.ps1 and continue-rip.ps1 (same pattern as
# Load-Config.ps1), so a resumed rip names episodes exactly like the original run
# would have. Genre series (-Series plus -Documentary etc.) does NOT use this file -
# it keeps its own numbering and move-up logic in Step 3 of each script.
#
# Final names (files stay in the per-disc DiscN folder, so no disc number is needed):
#   episodes: <Title>-S02-E05.mkv
#   extras:   <Title>-S02-Extra01.mkv
# With no -Season the season tag falls back to S01.

# ---------------------------------------------------------------------------
# Naming
# ---------------------------------------------------------------------------

function Get-SeriesSeasonTag {
    param([int]$Season)
    # No -Season: fall back to S01 rather than omit the tag, so every file matches the
    # one documented <Title>-S##-E## shape and Jellyfin always sees a season.
    if ($Season -gt 0) { return "S{0:D2}" -f $Season }
    return "S01"
}

function Get-SeriesEpisodeFileName {
    param([string]$Title, [int]$Season, [int]$Episode, [string]$Extension)
    return "{0}-{1}-E{2:D2}{3}" -f $Title, (Get-SeriesSeasonTag $Season), $Episode, $Extension
}

function Get-SeriesExtraFileName {
    param([string]$Title, [int]$Season, [int]$Extra, [string]$Extension)
    return "{0}-{1}-Extra{2:D2}{3}" -f $Title, (Get-SeriesSeasonTag $Season), $Extra, $Extension
}

# Regex matching a file already in the final shape for this title and season. Used to
# leave already-renamed files alone (e.g. re-running organize after a failure part-way
# through) and to reserve their numbers so nothing is ever numbered twice.
function Get-SeriesNamePattern {
    param([string]$Title, [int]$Season)
    $tag = Get-SeriesSeasonTag $Season
    return '^' + [regex]::Escape($Title) + '-' + $tag + '-(?:E(?<ep>\d+)|Extra(?<extra>\d+))\.(?:mp4|mkv)$'
}

# ---------------------------------------------------------------------------
# Starting episode (prompted at the start of the run for Disc 2+)
# ---------------------------------------------------------------------------

function Test-ShouldPromptStartEpisode {
    param(
        [switch]$Series,
        [switch]$GenreSeries,
        [switch]$Extras,
        [int]$Disc = 1,
        [switch]$StartEpisodeExplicit
    )
    # Only plain -Series, only Disc 2+, and only when the user has not already said
    # where numbering starts. Disc 1 keeps the old behaviour (start at -StartEpisode,
    # default 1) with no prompt; genre series has its own auto-detection.
    return [bool]($Series -and -not $GenreSeries -and -not $Extras -and $Disc -ge 2 -and -not $StartEpisodeExplicit)
}

# Highest episode number already on an EARLIER disc of this season, plus one. Reads
# both the renamed files and any rename-manifest.csv in sibling DiscN folders.
# Returns $null when nothing can be inferred.
function Get-SuggestedStartEpisode {
    param([string]$SeasonDir, [int]$Disc)

    if ([string]::IsNullOrWhiteSpace($SeasonDir) -or -not (Test-Path -LiteralPath $SeasonDir)) { return $null }

    $max = 0
    $discDirs = @(Get-ChildItem -LiteralPath $SeasonDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^Disc(\d+)$' -and [int]$Matches[1] -lt $Disc })

    foreach ($dir in $discDirs) {
        $names = @(Get-ChildItem -LiteralPath $dir.FullName -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
        $manifest = Join-Path $dir.FullName 'rename-manifest.csv'
        if (Test-Path -LiteralPath $manifest) {
            try {
                $names += @(Import-Csv -LiteralPath $manifest | Where-Object { $_.Kind -ne 'Extra' } | ForEach-Object { $_.NewName })
            } catch { }
        }
        foreach ($name in $names) {
            if ($name -match '-S\d+-E(\d+)\.(?:mp4|mkv)$') {
                $n = [int]$Matches[1]
                if ($n -gt $max) { $max = $n }
            }
        }
    }

    if ($max -gt 0) { return $max + 1 }
    return $null
}

function Read-StartEpisode {
    param(
        [int]$Disc,
        $Suggested = $null,
        # Injectable so tests can drive the prompt without a console.
        [scriptblock]$ReadInput = { param($p) Read-Host $p }
    )

    $prompt = if ($Suggested) { "Starting episode number for Disc $Disc [$Suggested]" } else { "Starting episode number for Disc $Disc" }
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        $answer = & $ReadInput $prompt
        if ($null -eq $answer) {
            # Non-interactive input: take the suggestion if there is one, otherwise stop
            # rather than spin forever.
            if ($Suggested) { return [int]$Suggested }
            throw "No input available for the starting episode number - pass -StartEpisode N on the command line."
        }
        $answer = "$answer".Trim()
        if ($answer -eq '' -and $Suggested) { return [int]$Suggested }
        if ($answer -match '^\d+$' -and [int]$answer -ge 1) { return [int]$answer }
        Write-Host "Please enter a whole number of 1 or more." -ForegroundColor Red
    }
    throw "No valid starting episode number entered - pass -StartEpisode N on the command line."
}

# ---------------------------------------------------------------------------
# Durations
# ---------------------------------------------------------------------------

function ConvertFrom-HandBrakeScanDuration {
    param([string[]]$Lines)
    foreach ($line in $Lines) {
        if ("$line" -match '\+\s*duration:\s*(\d+):(\d{2}):(\d{2})') {
            return ([int]$Matches[1] * 3600) + ([int]$Matches[2] * 60) + [int]$Matches[3]
        }
    }
    return $null
}

# Reads a video file's running time in seconds. The disc has already been ejected by
# Step 3, so durations come from the encoded files themselves (the encode keeps the
# title's running time). Windows' own property system first (fast, no extra tools),
# then a HandBrakeCLI scan. Returns $null when neither works.
function Get-VideoDurationSeconds {
    param([string]$Path, [string]$HandBrakePath = "")

    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Split-Path -Parent $Path))
        if ($folder) {
            $item = $folder.ParseName((Split-Path -Leaf $Path))
            if ($item) {
                # System.Media.Duration is in 100-nanosecond units and locale-independent.
                $ticks = $item.ExtendedProperty('System.Media.Duration')
                if ($ticks -and [double]$ticks -gt 0) { return [math]::Round([double]$ticks / 1e7) }
            }
        }
    } catch { }

    if ($HandBrakePath -and (Test-Path -LiteralPath $HandBrakePath)) {
        try {
            $scan = & $HandBrakePath --scan -i $Path 2>&1 | ForEach-Object { "$_" }
            $seconds = ConvertFrom-HandBrakeScanDuration -Lines $scan
            if ($seconds) { return $seconds }
        } catch { }
    }
    return $null
}

function Format-SeriesDuration {
    param($Seconds)
    if ($null -eq $Seconds -or $Seconds -le 0) { return "?" }
    $ts = [TimeSpan]::FromSeconds([double]$Seconds)
    if ($ts.TotalHours -ge 1) { return "{0}:{1:D2}:{2:D2}" -f [int][math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds }
    return "{0}:{1:D2}" -f $ts.Minutes, $ts.Seconds
}

function Get-Median {
    param([double[]]$Values)
    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) { return $null }
    $mid = [int][math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2 -eq 1) { return [double]$sorted[$mid] }
    return ([double]$sorted[$mid - 1] + [double]$sorted[$mid]) / 2
}

# ---------------------------------------------------------------------------
# Episode / extra classification
# ---------------------------------------------------------------------------

# Decides which titles on a disc are episodes and which are extras, and numbers them.
#
#   1. Play-all / composite: with 3+ titles, a title whose length is 70-130% of the sum
#      of the others is an extra (same threshold as the Step 2 composite check).
#   2. With TMDb season data: each title is compared with the published runtime of the
#      episode it would become (E<next>), within max(15%, 3 min). A match is an
#      episode. A title under 60% of that runtime is an extra and does not use up the
#      episode number. Anything else stays an episode but is flagged as a mismatch.
#   3. Without TMDb data (no key, no match, offline) or beyond the season's last
#      episode: the median heuristic - under 60% of the median title length is an extra.
#   4. A title with no readable duration stays an episode, flagged.
#
# -Overrides (OriginalName -> 'Episode'|'Extra') comes from the confirmation prompt's
# edit option and always wins. Numbers in -TakenEpisodes / -TakenExtras (files already
# renamed in this folder) are skipped so nothing is numbered twice.
function Get-SeriesTitleClassification {
    param(
        [object[]]$Titles = @(),
        [int]$StartEpisode = 1,
        [object[]]$TmdbEpisodes = @(),
        [int]$TmdbEpisodeCount = 0,
        [int[]]$TakenEpisodes = @(),
        [int[]]$TakenExtras = @(),
        [hashtable]$Overrides = @{},
        [switch]$AllExtras,
        [double]$ShortRatio = 0.6,
        [double]$TolerancePct = 0.15,
        [double]$ToleranceMinSec = 180
    )

    $items = @(foreach ($t in $Titles) {
        $d = if ($null -ne $t.DurationSec -and [double]$t.DurationSec -gt 0) { [double]$t.DurationSec } else { $null }
        [pscustomobject]@{
            Name          = $t.Name
            DurationSec   = $d
            Kind          = $null
            EpisodeNumber = $null
            ExtraNumber   = $null
            ExpectedSec   = $null
            Note          = ""
            Mismatch      = $false
        }
    })

    # --- 1. play-all / composite ---
    $known = @($items | Where-Object { $null -ne $_.DurationSec })
    if ($known.Count -ge 3) {
        $largest = $known | Sort-Object DurationSec -Descending | Select-Object -First 1
        $sumOthers = ($known | Where-Object { $_ -ne $largest } | Measure-Object -Property DurationSec -Sum).Sum
        if ($largest.DurationSec -ge ($sumOthers * 0.7) -and $largest.DurationSec -le ($sumOthers * 1.3)) {
            $largest.Kind = 'Extra'
            $largest.Note = 'play-all / composite (about the sum of the others)'
        }
    }

    $medianSec = Get-Median -Values @($known | Where-Object { $_.Kind -ne 'Extra' } | ForEach-Object { $_.DurationSec })

    $runtimeByEpisode = @{}
    foreach ($ep in $TmdbEpisodes) {
        if ($ep -and [int]$ep.Runtime -gt 0) { $runtimeByEpisode[[int]$ep.Number] = [int]$ep.Runtime * 60 }
    }
    $method = if ($runtimeByEpisode.Count -gt 0) { 'TMDb' } elseif ($null -ne $medianSec) { 'Median' } else { 'None' }

    $nextEpisode = [math]::Max(1, $StartEpisode)
    $nextExtra = 1
    $takenEp = @{}; foreach ($n in $TakenEpisodes) { $takenEp[[int]$n] = $true }
    $takenEx = @{}; foreach ($n in $TakenExtras) { $takenEx[[int]$n] = $true }

    foreach ($item in $items) {
        while ($takenEp.ContainsKey($nextEpisode)) { $nextEpisode++ }
        $expected = if ($runtimeByEpisode.ContainsKey($nextEpisode)) { $runtimeByEpisode[$nextEpisode] } else { $null }

        if ($Overrides -and $Overrides.ContainsKey($item.Name)) {
            $item.Kind = $Overrides[$item.Name]
            $item.Note = 'set by you'
        } elseif ($AllExtras) {
            $item.Kind = 'Extra'
            $item.Note = '-Extras disc'
        } elseif ($item.Kind -eq 'Extra') {
            # composite, already decided
        } elseif ($null -eq $item.DurationSec) {
            $item.Kind = 'Episode'
            $item.Note = 'duration unknown - check'
            $item.Mismatch = $true
        } elseif ($null -ne $expected) {
            $tolerance = [math]::Max($expected * $TolerancePct, $ToleranceMinSec)
            if ([math]::Abs($item.DurationSec - $expected) -le $tolerance) {
                $item.Kind = 'Episode'
                $item.Note = 'matches TMDb'
            } elseif ($item.DurationSec -lt ($expected * $ShortRatio)) {
                $item.Kind = 'Extra'
                $item.Note = "too short for E{0:D2} (TMDb ~{1})" -f $nextEpisode, (Format-SeriesDuration $expected)
            } else {
                $item.Kind = 'Episode'
                $item.Note = "differs from TMDb E{0:D2} runtime - check" -f $nextEpisode
                $item.Mismatch = $true
            }
        } else {
            if ($null -ne $medianSec -and $item.DurationSec -lt ($medianSec * $ShortRatio)) {
                $item.Kind = 'Extra'
                $item.Note = "short (under 60% of typical {0})" -f (Format-SeriesDuration $medianSec)
            } else {
                $item.Kind = 'Episode'
                if ($TmdbEpisodeCount -gt 0 -and $nextEpisode -gt $TmdbEpisodeCount) {
                    $item.Note = "past TMDb's $TmdbEpisodeCount episodes - check"
                    $item.Mismatch = $true
                }
            }
        }

        if ($item.Kind -eq 'Episode') {
            $item.EpisodeNumber = $nextEpisode
            $item.ExpectedSec = $expected
            $nextEpisode++
        } else {
            while ($takenEx.ContainsKey($nextExtra)) { $nextExtra++ }
            $item.ExtraNumber = $nextExtra
            $nextExtra++
        }
    }

    $warnings = @()
    $episodeItems = @($items | Where-Object { $_.Kind -eq 'Episode' })
    if ($TmdbEpisodeCount -gt 0 -and $episodeItems.Count -gt 0) {
        $remaining = [math]::Max(0, $TmdbEpisodeCount - [math]::Max(1, $StartEpisode) + 1)
        $lastNumber = ($episodeItems | Measure-Object -Property EpisodeNumber -Maximum).Maximum
        if ($episodeItems.Count -gt $remaining -or [int]$lastNumber -gt $TmdbEpisodeCount) {
            $warnings += "This disc has $($episodeItems.Count) episode-length title(s) but TMDb lists only $remaining episode(s) left in the season from E{0:D2} (season has $TmdbEpisodeCount) - numbering would run past the end of the season." -f [math]::Max(1, $StartEpisode)
        }
    }
    $mismatches = @($items | Where-Object { $_.Mismatch })
    if ($mismatches.Count -gt 0) {
        $warnings += "$($mismatches.Count) title(s) flagged for checking - see the Note column."
    }

    return [pscustomobject]@{
        Items     = $items
        Warnings  = $warnings
        Method    = $method
        MedianSec = $medianSec
    }
}

# Turns a classification into concrete renames. A target that already exists is
# marked Skip (never overwritten) and left out of the manifest.
function New-SeriesRenamePlan {
    param([object]$Classification, [string]$Title, [int]$Season, [string]$Directory)

    $plan = @(foreach ($item in $Classification.Items) {
        $ext = [System.IO.Path]::GetExtension($item.Name)
        $newName = if ($item.Kind -eq 'Episode') {
            Get-SeriesEpisodeFileName -Title $Title -Season $Season -Episode $item.EpisodeNumber -Extension $ext
        } else {
            Get-SeriesExtraFileName -Title $Title -Season $Season -Extra $item.ExtraNumber -Extension $ext
        }
        $newPath = Join-Path $Directory $newName
        $skip = Test-Path -LiteralPath $newPath
        [pscustomobject]@{
            OriginalName = $item.Name
            NewName      = $newName
            Kind         = $item.Kind
            DurationSec  = $item.DurationSec
            ExpectedSec  = $item.ExpectedSec
            Note         = if ($skip) { "TARGET EXISTS - will not overwrite" } else { $item.Note }
            OriginalPath = Join-Path $Directory $item.Name
            NewPath      = $newPath
            Skip         = $skip
        }
    })
    return ,$plan
}

function Show-SeriesRenamePlan {
    param([object[]]$Plan, [object]$Classification, [object]$TmdbSeason = $null)

    $methodText = switch ($Classification.Method) {
        'TMDb'   { "TMDb runtimes$(if ($TmdbSeason -and $TmdbSeason.ShowName) { " for $($TmdbSeason.ShowName)" })" }
        'Median' { "median title length ($(Format-SeriesDuration $Classification.MedianSec)) - no TMDb data" }
        default  { "no durations available" }
    }
    Write-Host "`nPlanned names (episodes vs extras by $methodText):" -ForegroundColor Cyan
    $nameWidth = [math]::Max(8, (($Plan | ForEach-Object { $_.OriginalName.Length } | Measure-Object -Maximum).Maximum))
    $newWidth = [math]::Max(8, (($Plan | ForEach-Object { $_.NewName.Length } | Measure-Object -Maximum).Maximum))
    Write-Host ("  {0,-3} {1} {2,8} {3,8}  {4,-7} {5} {6}" -f '#', 'Original'.PadRight($nameWidth), 'Actual', 'Expected', 'Kind', 'New name'.PadRight($newWidth), 'Note') -ForegroundColor Gray
    for ($i = 0; $i -lt $Plan.Count; $i++) {
        $p = $Plan[$i]
        $expected = if ($p.ExpectedSec) { Format-SeriesDuration $p.ExpectedSec } else { '-' }
        $color = if ($p.Skip) { 'Red' } elseif ($p.Kind -eq 'Extra') { 'DarkYellow' } elseif ($p.Note -match 'check') { 'Yellow' } else { 'White' }
        Write-Host ("  {0,-3} {1} {2,8} {3,8}  {4,-7} {5} {6}" -f ($i + 1), $p.OriginalName.PadRight($nameWidth), (Format-SeriesDuration $p.DurationSec), $expected, $p.Kind, $p.NewName.PadRight($newWidth), $p.Note) -ForegroundColor $color
    }
    foreach ($w in $Classification.Warnings) {
        Write-Host "  WARNING: $w" -ForegroundColor Yellow
    }
}

# Shows the plan and asks: Enter/Y accepts (the default), N leaves every file with its
# current name, E lets the user switch rows between episode and extra and re-plans.
# Returns @{ Classification; Plan } or $null when declined.
function Confirm-SeriesRenamePlan {
    param(
        [hashtable]$ClassifyArgs,
        [string]$Title,
        [int]$Season,
        [string]$Directory,
        [object]$TmdbSeason = $null,
        [scriptblock]$ReadInput = { param($p) Read-Host $p }
    )

    $overrides = @{}
    for ($round = 1; $round -le 50; $round++) {
        $classification = Get-SeriesTitleClassification @ClassifyArgs -Overrides $overrides
        $plan = New-SeriesRenamePlan -Classification $classification -Title $Title -Season $Season -Directory $Directory
        Show-SeriesRenamePlan -Plan $plan -Classification $classification -TmdbSeason $TmdbSeason

        $answer = & $ReadInput "Accept these names? [Y]es (default) / [n]o, leave files as they are / [e]dit"
        # No input available (redirected/non-interactive): take the default, accept.
        if ($null -eq $answer) { return @{ Classification = $classification; Plan = $plan } }

        switch -Regex ("$answer".Trim().ToLower()) {
            '^(|y|yes)$' { return @{ Classification = $classification; Plan = $plan } }
            '^(n|no)$'   { return $null }
            '^(e|edit)$' {
                $rows = & $ReadInput "Row number(s) to switch between episode and extra (e.g. 2 5)"
                foreach ($token in ("$rows" -split '[\s,]+' | Where-Object { $_ })) {
                    if ($token -match '^\d+$' -and [int]$token -ge 1 -and [int]$token -le $plan.Count) {
                        $row = $plan[[int]$token - 1]
                        $overrides[$row.OriginalName] = if ($row.Kind -eq 'Episode') { 'Extra' } else { 'Episode' }
                    } else {
                        Write-Host "  Ignoring '$token' - not a row number" -ForegroundColor Red
                    }
                }
            }
            default { Write-Host "Please answer Y, N or E." -ForegroundColor Red }
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Manifest and renaming
# ---------------------------------------------------------------------------

# Written BEFORE any file is renamed, so it survives a failure part-way through.
# Appends when a manifest already exists (e.g. organize re-run after a failure), so
# undo-rename.ps1 always sees every rename ever made in the folder.
function Write-RenameManifest {
    param([object[]]$Plan, [string]$ManifestPath)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $rows = @($Plan | Where-Object { -not $_.Skip } | ForEach-Object {
        [pscustomobject]@{
            OriginalName = $_.OriginalName
            NewName      = $_.NewName
            Kind         = $_.Kind
            OriginalPath = $_.OriginalPath
            NewPath      = $_.NewPath
            Timestamp    = $timestamp
        }
    })
    if ($rows.Count -eq 0) { return 0 }
    $append = Test-Path -LiteralPath $ManifestPath
    $rows | Export-Csv -LiteralPath $ManifestPath -NoTypeInformation -Encoding UTF8 -Append:$append
    return $rows.Count
}

# Renames each planned file. Never overwrites: a target that already exists is skipped
# with a warning. File-lock IOExceptions are retried, matching the rest of Step 3.
function Invoke-SeriesRenamePlan {
    param([object[]]$Plan, [int]$MaxRetries = 10, [int]$RetryDelaySec = 5)

    $renamed = 0
    $skipped = 0
    foreach ($p in $Plan) {
        if ($p.Skip -or (Test-Path -LiteralPath $p.NewPath)) {
            Write-Host "  SKIPPED $($p.OriginalName): $($p.NewName) already exists - not overwriting" -ForegroundColor Red
            Write-Log "WARNING: Not renaming $($p.OriginalName) - target $($p.NewName) already exists"
            $skipped++
            continue
        }
        if (-not (Test-Path -LiteralPath $p.OriginalPath)) {
            Write-Host "  SKIPPED $($p.OriginalName): file not found" -ForegroundColor Red
            Write-Log "WARNING: Not renaming $($p.OriginalName) - file not found"
            $skipped++
            continue
        }
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                Rename-Item -LiteralPath $p.OriginalPath -NewName $p.NewName -ErrorAction Stop
                break
            } catch [System.IO.IOException] {
                if ($attempt -eq $MaxRetries) {
                    Write-Host "  FAILED to rename $($p.OriginalName) after $MaxRetries attempts: $_" -ForegroundColor Red
                    Write-Log "ERROR: Failed to rename $($p.OriginalName) after $MaxRetries attempts: $_"
                    throw
                }
                Write-Host "  File locked: $($p.OriginalName) - retrying in ${RetryDelaySec}s (attempt $attempt/$MaxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelaySec
            }
        }
        Write-Host "  $($p.OriginalName) -> $($p.NewName)" -ForegroundColor Gray
        Write-Log "Renamed ($($p.Kind)): $($p.OriginalName) -> $($p.NewName)"
        $renamed++
    }
    return [pscustomobject]@{ Renamed = $renamed; Skipped = $skipped }
}

# Step 3 entry point for plain -Series. Works on the files in $Directory (the DiscN
# folder) and leaves them there.
function Invoke-SeriesEpisodeRename {
    param(
        [string]$Directory,
        [string]$Title,
        [int]$Season,
        [int]$StartEpisode = 1,
        [object]$TmdbSeason = $null,
        [switch]$AllExtras,
        [string]$UndoScriptSource = "",
        [string]$HandBrakePath = "",
        [scriptblock]$ReadInput = { param($p) Read-Host $p },
        # Injectable so tests do not need real video files.
        [scriptblock]$GetDuration = $null
    )

    $pattern = Get-SeriesNamePattern -Title $Title -Season $Season
    $videoFiles = @(Get-ChildItem -LiteralPath $Directory -File | Where-Object { $_.Extension -match '^\.(mp4|mkv)$' } | Sort-Object Name)

    $takenEpisodes = @()
    $takenExtras = @()
    $candidates = @()
    foreach ($f in $videoFiles) {
        if ($f.Name -match $pattern) {
            if ($Matches['ep']) { $takenEpisodes += [int]$Matches['ep'] } else { $takenExtras += [int]$Matches['extra'] }
        } else {
            $candidates += $f
        }
    }
    if ($takenEpisodes.Count + $takenExtras.Count -gt 0) {
        Write-Host "  $($takenEpisodes.Count + $takenExtras.Count) file(s) already renamed - leaving them alone and skipping their numbers" -ForegroundColor Gray
        Write-Log "Series rename: $($takenEpisodes.Count + $takenExtras.Count) file(s) already in final format, left unchanged"
    }
    if ($candidates.Count -eq 0) {
        Write-Host "  No files to rename" -ForegroundColor Gray
        Write-Log "Series rename: no files to rename in $Directory"
        return [pscustomobject]@{ Renamed = 0; Skipped = 0; Declined = $false; ManifestPath = $null }
    }

    Write-Host "  Reading durations of $($candidates.Count) file(s)..." -ForegroundColor Gray
    $titles = @(foreach ($f in $candidates) {
        $d = if ($GetDuration) { & $GetDuration $f.FullName } else { Get-VideoDurationSeconds -Path $f.FullName -HandBrakePath $HandBrakePath }
        [pscustomobject]@{ Name = $f.Name; DurationSec = $d }
    })

    $classifyArgs = @{
        Titles           = $titles
        StartEpisode     = $StartEpisode
        TmdbEpisodes     = if ($TmdbSeason) { @($TmdbSeason.Episodes) } else { @() }
        TmdbEpisodeCount = if ($TmdbSeason) { [int]$TmdbSeason.EpisodeCount } else { 0 }
        TakenEpisodes    = $takenEpisodes
        TakenExtras      = $takenExtras
        AllExtras        = [bool]$AllExtras
    }
    $result = Confirm-SeriesRenamePlan -ClassifyArgs $classifyArgs -Title $Title -Season $Season -Directory $Directory -TmdbSeason $TmdbSeason -ReadInput $ReadInput
    if (-not $result) {
        Write-Host "  Rename declined - files keep their current names. Re-run later with: continue-rip.ps1 ... -FromStep organize" -ForegroundColor Yellow
        Write-Log "Series rename: declined at the confirmation prompt - files left unchanged"
        return [pscustomobject]@{ Renamed = 0; Skipped = 0; Declined = $true; ManifestPath = $null }
    }

    $plan = @($result.Plan)
    Write-Log "Series rename: classification by $($result.Classification.Method); start episode $StartEpisode"
    foreach ($w in $result.Classification.Warnings) { Write-Log "Series rename WARNING: $w" }

    $manifestPath = Join-Path $Directory 'rename-manifest.csv'
    $rowCount = Write-RenameManifest -Plan $plan -ManifestPath $manifestPath
    Write-Host "  Rename manifest: $manifestPath ($rowCount row(s))" -ForegroundColor Gray
    Write-Log "Rename manifest written before renaming: $manifestPath ($rowCount row(s))"

    if ($UndoScriptSource -and (Test-Path -LiteralPath $UndoScriptSource)) {
        $undoTarget = Join-Path $Directory 'undo-rename.ps1'
        Copy-Item -LiteralPath $UndoScriptSource -Destination $undoTarget -Force
        Write-Log "Undo script copied: $undoTarget"
    } elseif ($UndoScriptSource) {
        Write-Host "  WARNING: undo script not found at $UndoScriptSource - use the repo's undo-rename.ps1 -ManifestPath `"$manifestPath`"" -ForegroundColor Yellow
        Write-Log "WARNING: undo script not found at $UndoScriptSource"
    }

    $outcome = Invoke-SeriesRenamePlan -Plan $plan
    $episodes = @($plan | Where-Object { $_.Kind -eq 'Episode' -and -not $_.Skip }).Count
    $extras = @($plan | Where-Object { $_.Kind -eq 'Extra' -and -not $_.Skip }).Count
    Write-Host "Renamed $($outcome.Renamed) file(s) ($episodes episode(s), $extras extra(s))$(if ($outcome.Skipped) { ", skipped $($outcome.Skipped)" })" -ForegroundColor Green
    Write-Host "  To undo: & `"$(Join-Path $Directory 'undo-rename.ps1')`"   (add -WhatIf to preview)" -ForegroundColor Gray
    Write-Log "Series rename: $($outcome.Renamed) renamed, $($outcome.Skipped) skipped"

    return [pscustomobject]@{ Renamed = $outcome.Renamed; Skipped = $outcome.Skipped; Declined = $false; ManifestPath = $manifestPath }
}

# ---------------------------------------------------------------------------
# TMDb season lookup (optional, fail-soft)
# ---------------------------------------------------------------------------

# Single seam for every TMDb HTTP call - tests replace this function with a mock.
function Invoke-TMDbRequest {
    param([string]$Url)
    return Invoke-RestMethod -Uri $Url -Method Get -TimeoutSec 10
}

function Get-TMDbApiKey {
    if ($script:Config_TmdbApiKey) { return $script:Config_TmdbApiKey }
    if ($env:TMDB_API_KEY) { return $env:TMDB_API_KEY }
    return $null
}

# Non-interactive TV search: exact (case-insensitive) name match wins, otherwise the
# top result. The match is printed so a wrong pick is visible before the rip starts.
function Find-TMDbTvShow {
    param([string]$Title, [string]$ApiKey)
    $url = "https://api.themoviedb.org/3/search/tv?query={0}&api_key={1}" -f [System.Uri]::EscapeDataString($Title), $ApiKey
    $response = Invoke-TMDbRequest -Url $url
    $results = @($response.results)
    if ($results.Count -eq 0 -or $null -eq $results[0]) { return $null }
    $pick = $results | Where-Object { $_.name -and $_.name.Trim() -ieq $Title.Trim() } | Select-Object -First 1
    if (-not $pick) { $pick = $results[0] }
    return [pscustomobject]@{
        Id   = [int]$pick.id
        Name = $pick.name
        Year = if ($pick.first_air_date) { ($pick.first_air_date -split '-')[0] } else { "" }
    }
}

function Get-TMDbSeasonDetails {
    param([int]$TvId, [int]$Season, [string]$ApiKey)
    $url = "https://api.themoviedb.org/3/tv/{0}/season/{1}?api_key={2}" -f $TvId, $Season, $ApiKey
    $response = Invoke-TMDbRequest -Url $url
    $episodes = @(foreach ($e in @($response.episodes)) {
        if ($null -eq $e) { continue }
        [pscustomobject]@{
            Number  = [int]$e.episode_number
            Name    = $e.name
            Runtime = if ($e.runtime) { [int]$e.runtime } else { 0 }
        }
    })
    return [pscustomobject]@{
        TvId         = $TvId
        Season       = $Season
        EpisodeCount = $episodes.Count
        Episodes     = $episodes
        ShowName     = ""
    }
}

# Looks up the season's per-episode runtimes. Every failure (no key, no match, offline,
# bad response) returns $null with a one-line note, and classification falls back to
# the median heuristic - this lookup never stops a rip.
function Get-SeriesTmdbSeason {
    param([string]$Title, [int]$Season, $KnownTvId = $null)

    $apiKey = Get-TMDbApiKey
    if (-not $apiKey) {
        Write-Host "TMDb key not set - episode/extra detection will use title lengths only" -ForegroundColor DarkGray
        return $null
    }
    $tmdbSeasonNumber = if ($Season -gt 0) { $Season } else { 1 }
    try {
        $showName = $Title
        $tvId = $KnownTvId
        if (-not $tvId) {
            $show = Find-TMDbTvShow -Title $Title -ApiKey $apiKey
            if (-not $show) {
                Write-Host "TMDb: no TV match for '$Title' - episode/extra detection will use title lengths only" -ForegroundColor DarkGray
                return $null
            }
            $tvId = $show.Id
            $showName = "$($show.Name)$(if ($show.Year) { " ($($show.Year))" })"
        }
        $details = Get-TMDbSeasonDetails -TvId $tvId -Season $tmdbSeasonNumber -ApiKey $apiKey
        if (-not $details -or $details.EpisodeCount -eq 0) {
            Write-Host "TMDb: no episode data for season $tmdbSeasonNumber of $showName - using title lengths only" -ForegroundColor DarkGray
            return $null
        }
        $details.ShowName = $showName
        Write-Host "TMDb: $showName, season $tmdbSeasonNumber - $($details.EpisodeCount) episode(s)" -ForegroundColor Gray
        return $details
    } catch {
        Write-Host "TMDb season lookup failed ($($_.Exception.Message)) - using title lengths only" -ForegroundColor DarkGray
        return $null
    }
}
