# SeriesEpisodes.ps1 - plain -Series episode numbering, extras classification,
# rename manifest and optional TMDb season lookup.
#
# Dot-sourced by both rip-disc.ps1 and continue-rip.ps1 (same pattern as
# Load-Config.ps1), so a resumed rip names episodes exactly like the original run
# would have. Genre series (-Series plus -Documentary etc.) does NOT use this file -
# it keeps its own numbering and move-up logic in Step 3 of each script.
#
# Final names (episodes stay in the per-disc DiscN folder):
#   episodes: <Series>\Season 2\Disc1\<Title>-S02-E05.mkv
#   extras:   <Series>\Specials\<Title>-S02-D1-Extra01.mkv   (or -Extra01-<TheDiscDB name>)
# With no -Season the season tag falls back to S01.
#
# Extras go in ONE "Specials" folder at series level, alongside the Season folders -
# Jellyfin did not pick up the earlier DiscN\extras\ (PR #143), and treats a series-level
# Specials folder as Season 00. Every disc of every season shares that folder, so an
# extra's name carries its season AND disc (-S02-D1-): Extra## is numbered per disc, and
# the disc part keeps two discs' -Extra01 apart. Concurrent rips of different discs only
# ever create distinct names there, and renames never overwrite. The manifest stays in
# the DiscN folder, recording each extra's path relative to it (..\..\Specials\<name>).

# Series-level folder for extras. Must match undo-rename.ps1's NewName whitelist.
$script:SeriesExtrasFolder = 'Specials'
# Where PR #143 put extras (DiscN\extras\). Still read so extras moved there by an older
# run keep their numbers, and still accepted by undo-rename.ps1 for older manifests.
$script:SeriesLegacyExtrasFolder = 'extras'

# The Specials folder for a folder of episodes, as a path RELATIVE to it: up past DiscN,
# then up past Season N, to the series folder. "..\..\Specials" for Season N\DiscN,
# "..\Specials" for a DiscN or Season N folder directly under the series.
function Get-SeriesSpecialsRelativeDir {
    param([string]$Directory)
    $dir = $Directory.TrimEnd('\', '/')
    $ups = 0
    if ((Split-Path $dir -Leaf) -match '^Disc\d+$') { $dir = Split-Path $dir -Parent; $ups++ }
    if ($dir -and (Split-Path $dir -Leaf) -match '^Season\s+\d+$') { $ups++ }
    $prefix = ('..\' * $ups)
    return "$prefix$($script:SeriesExtrasFolder)"
}

# Disc number from a DiscN folder name, 0 when the folder is not a DiscN folder (a Season
# folder that holds the files itself) - the name then has no -D# part.
function Get-SeriesDiscFromDirectory {
    param([string]$Directory)
    if ((Split-Path $Directory.TrimEnd('\', '/') -Leaf) -match '^Disc(\d+)$') { return [int]$Matches[1] }
    return 0
}

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
    param([string]$Title, [int]$Season, [int]$Extra, [string]$Extension, [string]$Label = "", [int]$Disc = 0)
    # -Label is TheDiscDB's published name for the extra (already made filename-safe by
    # Get-SeriesExtraLabel). It goes AFTER the number, so sorting and uniqueness are
    # still decided by D#-Extra## alone: <Title>-S02-D1-Extra01-Making Of.mkv
    # -Disc 0 (files not in a DiscN folder) leaves the -D# part out.
    $suffix = if ($Label) { "-$Label" } else { "" }
    $discPart = if ($Disc -gt 0) { "D$Disc-" } else { "" }
    return "{0}-{1}-{2}Extra{3:D2}{4}{5}" -f $Title, (Get-SeriesSeasonTag $Season), $discPart, $Extra, $suffix, $Extension
}

# Turns a published extra name into something safe to put in a filename: characters
# Windows forbids (\ / : * ? " < > |) and control characters become spaces, runs of
# whitespace collapse, leading/trailing spaces, dots and hyphens go (Windows silently
# strips trailing dots/spaces), and the result is capped at -MaxLength characters,
# cut at a word boundary where possible. Returns "" when nothing usable is left.
function Get-SeriesExtraLabel {
    param([string]$Name, [int]$MaxLength = 50)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    $label = $Name -replace '[\\/:*?"<>|\x00-\x1F\x7F]', ' '
    $label = ($label -replace '\s+', ' ').Trim(' ', '.', '-')
    if ($label.Length -gt $MaxLength) {
        $cut = $label.Substring(0, $MaxLength)
        $lastSpace = $cut.LastIndexOf(' ')
        if ($lastSpace -ge [int]($MaxLength / 2)) { $cut = $cut.Substring(0, $lastSpace) }
        $label = $cut.Trim(' ', '.', '-')
    }
    return $label
}

# Regex matching a file already in the final shape for this title and season. Used to
# leave already-renamed files alone (e.g. re-running organize after a failure part-way
# through) and to reserve their numbers so nothing is ever numbered twice. Extras may
# carry a disc part (-D2-, absent on PR #143-era names) and a TheDiscDB name after the
# number (-Extra01-Making Of).
function Get-SeriesNamePattern {
    param([string]$Title, [int]$Season)
    $tag = Get-SeriesSeasonTag $Season
    return '^' + [regex]::Escape($Title) + '-' + $tag + '-(?:E(?<ep>\d+)|(?:D(?<disc>\d+)-)?Extra(?<extra>\d+)(?:-[^\\/]+)?)\.(?:mp4|mkv)$'
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

# Non-interactive answer to the start-episode question (continue-rip.ps1 -Yes): the
# suggested default when there is one (TheDiscDB's first episode, or the next number
# after earlier discs), otherwise -Fallback (the current -StartEpisode, normally 1) -
# a non-interactive run must not stop to ask. Returns the number and a one-line reason
# for the log, flagged when it is only the fallback.
function Get-AutoStartEpisode {
    param($Suggested = $null, [int]$Fallback = 1, [string]$SuggestedFrom = "suggested default")
    if ($Suggested -and [int]$Suggested -ge 1) {
        return [pscustomobject]@{ Episode = [int]$Suggested; Reason = "$SuggestedFrom (-Yes, not asked)"; IsGuess = $false }
    }
    $n = [math]::Max(1, $Fallback)
    return [pscustomobject]@{ Episode = $n; Reason = "no suggestion available - defaulted to E{0:D2} (-Yes, not asked) - check the names" -f $n; IsGuess = $true }
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
# TheDiscDB comes before all of that: -DiscDbMap (OriginalName -> TheDiscDB title, from
# Get-DiscDbTitleMatches) gives a file its published episode number, or marks it as a
# named extra. Its episode numbers are reserved up front so the sequential numbering of
# any file TheDiscDB does not cover skips them. A TheDiscDB episode number that is
# already taken falls back to the steps above instead of being used twice.
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
        [hashtable]$DiscDbMap = @{},
        # The season the files are being named for - a TheDiscDB title from a different
        # season is still used, but flagged for checking.
        [int]$Season = 0,
        [double]$ShortRatio = 0.6,
        [double]$TolerancePct = 0.15,
        [double]$ToleranceMinSec = 180
    )
    if (-not $DiscDbMap) { $DiscDbMap = @{} }
    if (-not $Overrides) { $Overrides = @{} }

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
            # Where the episode/extra decision came from: TheDiscDB, TMDb, Length
            # (median / play-all heuristics), You (edited at the prompt), -Extras, or
            # '-' (no duration to go on).
            Source        = ""
            # TheDiscDB's published name for an extra, filename-safe; "" otherwise.
            Label         = ""
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

    # --- 0. TheDiscDB episode numbers, reserved before any sequential numbering ---
    $discDbEpisode = @{}
    if ($DiscDbMap.Count -gt 0 -and -not $AllExtras) {
        $reserved = @{}
        foreach ($item in $items) {
            if ($Overrides.ContainsKey($item.Name)) { continue }
            $dd = $DiscDbMap[$item.Name]
            if (-not $dd -or $dd.ItemType -ne 'Episode' -or $null -eq $dd.Episode -or [int]$dd.Episode -lt 1) { continue }
            $n = [int]$dd.Episode
            if ($takenEp.ContainsKey($n) -or $reserved.ContainsKey($n)) { continue }
            $discDbEpisode[$item.Name] = $n
            $reserved[$n] = $true
        }
        foreach ($n in $reserved.Keys) { $takenEp[$n] = $true }
    }

    foreach ($item in $items) {
        while ($takenEp.ContainsKey($nextEpisode)) { $nextEpisode++ }
        $expected = if ($runtimeByEpisode.ContainsKey($nextEpisode)) { $runtimeByEpisode[$nextEpisode] } else { $null }
        $dd = $DiscDbMap[$item.Name]
        $ddExtraLabel = if ($dd -and $dd.ItemType -and $dd.ItemType -ne 'Episode') { Get-SeriesExtraLabel $dd.Name } else { "" }

        if ($Overrides.ContainsKey($item.Name)) {
            $item.Kind = $Overrides[$item.Name]
            $item.Note = 'set by you'
            $item.Source = 'You'
            if ($item.Kind -eq 'Extra') { $item.Label = $ddExtraLabel }
        } elseif ($AllExtras) {
            $item.Kind = 'Extra'
            $item.Note = '-Extras disc'
            $item.Source = '-Extras'
            $item.Label = $ddExtraLabel
        } elseif ($discDbEpisode.ContainsKey($item.Name)) {
            $item.Kind = 'Episode'
            $item.Source = 'TheDiscDB'
            $item.Note = if ($dd.Name) { "TheDiscDB: $($dd.Name)" } else { 'TheDiscDB' }
            if ($Season -gt 0 -and $null -ne $dd.Season -and [int]$dd.Season -ne $Season) {
                $item.Note = "TheDiscDB lists this as S{0:D2}E{1:D2} - check" -f [int]$dd.Season, [int]$dd.Episode
                $item.Mismatch = $true
            }
        } elseif ($dd -and $dd.ItemType -and $dd.ItemType -ne 'Episode') {
            $item.Kind = 'Extra'
            $item.Source = 'TheDiscDB'
            $item.Label = $ddExtraLabel
            $item.Note = "TheDiscDB: $($dd.ItemType)"
        } elseif ($item.Kind -eq 'Extra') {
            # composite, already decided
            $item.Source = 'Length'
        } elseif ($null -eq $item.DurationSec) {
            $item.Kind = 'Episode'
            $item.Note = 'duration unknown - check'
            $item.Mismatch = $true
            $item.Source = '-'
        } elseif ($null -ne $expected) {
            $item.Source = 'TMDb'
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
            $item.Source = 'Length'
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

        # The disc WAS found in TheDiscDB but this episode-length file matched none of
        # its titles - the sequential number it gets is a guess, so say so.
        if ($DiscDbMap.Count -gt 0 -and $item.Kind -eq 'Episode' -and $item.Source -in @('TMDb', 'Length') -and -not $item.Mismatch) {
            $item.Note = (@($item.Note, 'not matched in TheDiscDB - check') | Where-Object { $_ }) -join '; '
            $item.Mismatch = $true
        }

        if ($item.Kind -eq 'Episode' -and $item.Source -eq 'TheDiscDB') {
            # Published number - does not use up the sequential counter.
            $item.EpisodeNumber = $discDbEpisode[$item.Name]
            $item.ExpectedSec = $dd.DurationSec
        } elseif ($item.Kind -eq 'Episode') {
            $item.EpisodeNumber = $nextEpisode
            $item.ExpectedSec = $expected
            $nextEpisode++
        } else {
            if ($item.Source -eq 'TheDiscDB') { $item.ExpectedSec = $dd.DurationSec }
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

    $discDbCount = @($items | Where-Object { $_.Source -eq 'TheDiscDB' }).Count
    return [pscustomobject]@{
        Items        = $items
        Warnings     = $warnings
        # The fallback used for files TheDiscDB does not cover (TMDb / Median / None).
        Method       = $method
        MedianSec    = $medianSec
        DiscDbCount  = $discDbCount
    }
}

# Turns a classification into concrete renames. A target that already exists is
# marked Skip (never overwritten) and left out of the manifest.
function New-SeriesRenamePlan {
    param([object]$Classification, [string]$Title, [int]$Season, [string]$Directory)

    $specialsRel = Get-SeriesSpecialsRelativeDir -Directory $Directory
    $disc = Get-SeriesDiscFromDirectory -Directory $Directory
    $plan = @(foreach ($item in $Classification.Items) {
        $ext = [System.IO.Path]::GetExtension($item.Name)
        # NewName is the path RELATIVE to $Directory (what the manifest records and
        # undo-rename.ps1 resolves): a bare name for episodes, ..\..\Specials\<name> for extras.
        if ($item.Kind -eq 'Episode') {
            $fileName = Get-SeriesEpisodeFileName -Title $Title -Season $Season -Episode $item.EpisodeNumber -Extension $ext
            $newName = $fileName
        } else {
            $fileName = Get-SeriesExtraFileName -Title $Title -Season $Season -Extra $item.ExtraNumber -Extension $ext -Label $item.Label -Disc $disc
            $newName = "$specialsRel\$fileName"
        }
        # GetFullPath folds the ..\ segments so NewPath (manifest, messages) is a clean path.
        $newPath = [System.IO.Path]::GetFullPath((Join-Path $Directory $newName))
        $skip = Test-Path -LiteralPath $newPath
        [pscustomobject]@{
            OriginalName = $item.Name
            NewName      = $newName
            FileName     = $fileName
            Kind         = $item.Kind
            Source       = $item.Source
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
    param([object[]]$Plan, [object]$Classification, [object]$TmdbSeason = $null, [object]$DiscDbDisc = $null)

    $methodText = switch ($Classification.Method) {
        'TMDb'   { "TMDb runtimes$(if ($TmdbSeason -and $TmdbSeason.ShowName) { " for $($TmdbSeason.ShowName)" })" }
        'Median' { "median title length ($(Format-SeriesDuration $Classification.MedianSec)) - no TMDb data" }
        default  { "no durations available" }
    }
    if ($Classification.DiscDbCount -gt 0) {
        $discText = if ($DiscDbDisc -and $DiscDbDisc.Description) { " ($($DiscDbDisc.Description))" } else { "" }
        $rest = $Plan.Count - $Classification.DiscDbCount
        $methodText = "TheDiscDB$discText for $($Classification.DiscDbCount) of $($Plan.Count) title(s)$(if ($rest -gt 0) { "; the rest by $methodText" })"
    }
    Write-Host "`nPlanned names (episodes vs extras by $methodText):" -ForegroundColor Cyan
    $nameWidth = [math]::Max(8, (($Plan | ForEach-Object { $_.OriginalName.Length } | Measure-Object -Maximum).Maximum))
    $newWidth = [math]::Max(8, (($Plan | ForEach-Object { $_.NewName.Length } | Measure-Object -Maximum).Maximum))
    $row = "  {0,-3} {1} {2,8} {3,8}  {4,-7} {5,-9} {6} {7}"
    Write-Host ($row -f '#', 'Original'.PadRight($nameWidth), 'Actual', 'Expected', 'Kind', 'Source', 'New name'.PadRight($newWidth), 'Note') -ForegroundColor Gray
    for ($i = 0; $i -lt $Plan.Count; $i++) {
        $p = $Plan[$i]
        $expected = if ($p.ExpectedSec) { Format-SeriesDuration $p.ExpectedSec } else { '-' }
        $color = if ($p.Skip) { 'Red' } elseif ($p.Kind -eq 'Extra') { 'DarkYellow' } elseif ($p.Note -match 'check') { 'Yellow' } else { 'White' }
        Write-Host ($row -f ($i + 1), $p.OriginalName.PadRight($nameWidth), (Format-SeriesDuration $p.DurationSec), $expected, $p.Kind, $p.Source, $p.NewName.PadRight($newWidth), $p.Note) -ForegroundColor $color
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
        [object]$DiscDbDisc = $null,
        [scriptblock]$ReadInput = { param($p) Read-Host $p },
        # Non-interactive (continue-rip.ps1 -Yes): show the table, then accept it as-is.
        [switch]$AutoAccept
    )

    $overrides = @{}
    for ($round = 1; $round -le 50; $round++) {
        $classification = Get-SeriesTitleClassification @ClassifyArgs -Overrides $overrides
        $plan = New-SeriesRenamePlan -Classification $classification -Title $Title -Season $Season -Directory $Directory
        Show-SeriesRenamePlan -Plan $plan -Classification $classification -TmdbSeason $TmdbSeason -DiscDbDisc $DiscDbDisc

        if ($AutoAccept) {
            Write-Host "Accepted automatically (-Yes) - undo-rename.ps1 reverses it if it is wrong." -ForegroundColor Gray
            return @{ Classification = $classification; Plan = $plan; AutoAccepted = $true }
        }

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
        # Extras move into the series-level Specials folder; create it on first use only,
        # so a series with no extras gets no empty folder.
        $targetDir = Split-Path -Parent $p.NewPath
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                # Move-Item, not Rename-Item, so extras can change folder. Same volume, so
                # it is still a rename on disk; it never overwrites (no -Force).
                Move-Item -LiteralPath $p.OriginalPath -Destination $p.NewPath -ErrorAction Stop
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
        # TheDiscDB disc record from Get-SeriesDiscDbLookup (captured while the disc was
        # still in the drive). $null = not found / offline / -NoDiscDb.
        [object]$DiscDbDisc = $null,
        [switch]$AllExtras,
        [string]$UndoScriptSource = "",
        [string]$HandBrakePath = "",
        [scriptblock]$ReadInput = { param($p) Read-Host $p },
        # Injectable so tests do not need real video files.
        [scriptblock]$GetDuration = $null,
        # Accept the confirmation table without asking (continue-rip.ps1 -Yes).
        [switch]$AutoAccept,
        # Preview only (rename-series.ps1 without -Apply): classify and show the planned
        # names, then stop - no prompt, no manifest, no undo script, no file touched.
        # The returned Plan is what an apply run would do.
        [switch]$DryRun,
        # Only files whose NAME this accepts are renamed (rename-series.ps1: several discs'
        # files sharing one folder, grouped by a "Disc N" token in the file name). Files
        # already in final shape still reserve their numbers whatever the filter says.
        [scriptblock]$FileFilter = $null,
        # Extra numbers already planned for this folder by an earlier unit in the same run
        # (a dry run has not renamed them yet), so two units sharing a folder never plan
        # the same -Extra##.
        [int[]]$ReservedExtras = @()
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
            if (-not $FileFilter -or (& $FileFilter $f.Name)) { $candidates += $f }
        }
    }
    # Extras already moved out by an earlier run keep their numbers: this disc's own
    # (-D#-) extras in the shared Specials folder, and anything in a PR #143-era
    # DiscN\extras\ folder (those names have no disc part and belong to this disc).
    $disc = Get-SeriesDiscFromDirectory -Directory $Directory
    $extrasDir = [System.IO.Path]::GetFullPath((Join-Path $Directory (Get-SeriesSpecialsRelativeDir -Directory $Directory)))
    $legacyExtrasDir = Join-Path $Directory $script:SeriesLegacyExtrasFolder
    foreach ($dir in @($extrasDir, $legacyExtrasDir)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Extension -match '^\.(mp4|mkv)$' })) {
            if (-not ($f.Name -match $pattern) -or -not $Matches['extra']) { continue }
            $fileDisc = if ($Matches['disc']) { [int]$Matches['disc'] } else { 0 }
            if ($dir -eq $legacyExtrasDir -or $fileDisc -eq $disc) { $takenExtras += [int]$Matches['extra'] }
        }
    }
    if ($takenEpisodes.Count + $takenExtras.Count -gt 0) {
        Write-Host "  $($takenEpisodes.Count + $takenExtras.Count) file(s) already renamed - leaving them alone and skipping their numbers" -ForegroundColor Gray
        Write-Log "Series rename: $($takenEpisodes.Count + $takenExtras.Count) file(s) already in final format, left unchanged"
    }
    # Reserved only now, so the "already renamed" count above stays about real files.
    $takenExtras += @($ReservedExtras)
    if ($candidates.Count -eq 0) {
        Write-Host "  No files to rename" -ForegroundColor Gray
        Write-Log "Series rename: no files to rename in $Directory"
        return [pscustomobject]@{ Renamed = 0; Skipped = 0; Declined = $false; ManifestPath = $null; Plan = @(); DryRun = [bool]$DryRun }
    }

    Write-Host "  Reading durations of $($candidates.Count) file(s)..." -ForegroundColor Gray
    $titles = @(foreach ($f in $candidates) {
        $d = if ($GetDuration) { & $GetDuration $f.FullName } else { Get-VideoDurationSeconds -Path $f.FullName -HandBrakePath $HandBrakePath }
        [pscustomobject]@{ Name = $f.Name; DurationSec = $d }
    })

    # TheDiscDB first: line the files up with the disc's published titles. Any problem
    # here just means no TheDiscDB rows - classification carries on with TMDb / median.
    $discDbMap = @{}
    if ($DiscDbDisc) {
        try {
            $discDbMap = Get-DiscDbTitleMatches -Titles $titles -DiscDbDisc $DiscDbDisc
        } catch {
            $discDbMap = @{}
            Write-Host "  TheDiscDB matching failed ($($_.Exception.Message)) - using TMDb / title lengths" -ForegroundColor DarkGray
        }
        $identified = @($discDbMap.Values | Where-Object { $_.ItemType }).Count
        Write-Log "TheDiscDB: $($discDbMap.Count) of $($titles.Count) file(s) lined up with $($DiscDbDisc.Description) ($identified identified)"
        if ($discDbMap.Count -eq 0) {
            Write-Host "  TheDiscDB: no file matched the disc's published title lengths - using TMDb / title lengths" -ForegroundColor DarkGray
        }
    }

    $classifyArgs = @{
        Titles           = $titles
        StartEpisode     = $StartEpisode
        TmdbEpisodes     = if ($TmdbSeason) { @($TmdbSeason.Episodes) } else { @() }
        TmdbEpisodeCount = if ($TmdbSeason) { [int]$TmdbSeason.EpisodeCount } else { 0 }
        TakenEpisodes    = $takenEpisodes
        TakenExtras      = $takenExtras
        AllExtras        = [bool]$AllExtras
        DiscDbMap        = $discDbMap
        Season           = $Season
    }
    if ($DryRun) {
        $classification = Get-SeriesTitleClassification @classifyArgs
        $plan = New-SeriesRenamePlan -Classification $classification -Title $Title -Season $Season -Directory $Directory
        Show-SeriesRenamePlan -Plan $plan -Classification $classification -TmdbSeason $TmdbSeason -DiscDbDisc $DiscDbDisc
        return [pscustomobject]@{ Renamed = 0; Skipped = 0; Declined = $false; ManifestPath = $null; Plan = @($plan); DryRun = $true }
    }

    $result = Confirm-SeriesRenamePlan -ClassifyArgs $classifyArgs -Title $Title -Season $Season -Directory $Directory -TmdbSeason $TmdbSeason -DiscDbDisc $DiscDbDisc -ReadInput $ReadInput -AutoAccept:$AutoAccept
    if (-not $result) {
        Write-Host "  Rename declined - files keep their current names. Re-run later with: continue-rip.ps1 ... -FromStep organize" -ForegroundColor Yellow
        Write-Log "Series rename: declined at the confirmation prompt - files left unchanged"
        return [pscustomobject]@{ Renamed = 0; Skipped = 0; Declined = $true; ManifestPath = $null; Plan = @(); DryRun = $false }
    }

    $plan = @($result.Plan)
    if ($result.AutoAccepted) {
        $autoEpisodes = @($plan | Where-Object { $_.Kind -eq 'Episode' }).Count
        Write-Log "Series rename: confirmation table accepted automatically (-Yes) - $autoEpisodes episode(s), $($plan.Count - $autoEpisodes) extra(s) as planned"
    }
    Write-Log "Series rename: classification by $(if ($result.Classification.DiscDbCount -gt 0) { "TheDiscDB ($($result.Classification.DiscDbCount) title(s)), then " })$($result.Classification.Method); start episode $StartEpisode"
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
    if ($extras -gt 0) { Write-Host "  Extras moved to: $extrasDir" -ForegroundColor Gray }
    Write-Host "  To undo: & `"$(Join-Path $Directory 'undo-rename.ps1')`"   (add -WhatIf to preview)" -ForegroundColor Gray
    Write-Log "Series rename: $($outcome.Renamed) renamed, $($outcome.Skipped) skipped"

    return [pscustomobject]@{ Renamed = $outcome.Renamed; Skipped = $outcome.Skipped; Declined = $false; ManifestPath = $manifestPath; Plan = $plan; DryRun = $false }
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

# ---------------------------------------------------------------------------
# TheDiscDB disc lookup (optional, fail-soft) - https://thediscdb.com
# ---------------------------------------------------------------------------
#
# TheDiscDB maps each MakeMKV title on a known disc to an episode (season/episode) or a
# named extra. It is keyed by a "content hash" of the disc's file sizes, so the disc is
# identified exactly - not by guessing from its name - and no API key is needed.
#
# Content hash (TheDiscDb/web: DiscScanner.cs + HashingExtensions.cs):
#   Blu-ray / UHD: every *.m2ts directly inside BDMV\STREAM
#   DVD:           every file directly inside VIDEO_TS
#   files ordered by name; MD5 over each file's size as a little-endian Int64;
#   uppercase hex with no separators (32 characters).
# Only the directory listing is read - no file contents - so it is quick, but it needs
# the disc in the drive. rip-disc.ps1 therefore computes it BEFORE the rip (the disc is
# ejected after Step 1) and hands the hash to continue-rip.ps1 as -DiscDbHash.

$script:DiscDbEndpoint = 'https://thediscdb.com/graphql'

# Pure hash over a list of sizes, already in TheDiscDB's order. Separate from the disc
# reading so it can be tested against known values.
function Get-DiscDbContentHashFromSizes {
    param([long[]]$Sizes)
    $bytes = New-Object System.Collections.Generic.List[byte]
    foreach ($s in $Sizes) { $bytes.AddRange([BitConverter]::GetBytes([long]$s)) }
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $digest = $md5.ComputeHash($bytes.ToArray())
    } finally {
        $md5.Dispose()
    }
    return ([BitConverter]::ToString($digest) -replace '-', '')
}

# Reads the disc's file listing and returns @{ Hash; Format; FileCount }, or $null when
# the drive has no BDMV\STREAM or VIDEO_TS (not a video disc, or not readable).
function Get-DiscDbContentHash {
    param([string]$DriveRoot)
    if ([string]::IsNullOrWhiteSpace($DriveRoot)) { return $null }
    $root = $DriveRoot.TrimEnd('\') + '\'

    $files = @()
    $format = $null
    $stream = Join-Path $root 'BDMV\STREAM'
    if (Test-Path -LiteralPath $stream -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $stream -File -Force -ErrorAction Stop | Where-Object { $_.Extension -eq '.m2ts' })
        $format = 'Blu-ray'
    }
    if ($files.Count -eq 0) {
        # TheDiscDB does not fall back to VIDEO_TS on a disc that has BDMV or AACS.
        if ((Test-Path -LiteralPath (Join-Path $root 'BDMV')) -or (Test-Path -LiteralPath (Join-Path $root 'AACS'))) { return $null }
        $videoTs = Join-Path $root 'VIDEO_TS'
        if (Test-Path -LiteralPath $videoTs -PathType Container) {
            $files = @(Get-ChildItem -LiteralPath $videoTs -File -Force -ErrorAction Stop)
            $format = 'DVD'
        }
    }
    if ($files.Count -eq 0) { return $null }

    # TheDiscDB sorts with .NET's default (culture-aware) string comparer; on-disc names
    # (00001.m2ts, VTS_01_1.VOB) sort the same under any culture, invariant used here.
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) { $list.Add($f) }
    $list.Sort([System.Comparison[object]] { param($a, $b) [string]::Compare($a.Name, $b.Name, [System.StringComparison]::InvariantCulture) })

    return [pscustomobject]@{
        Hash      = Get-DiscDbContentHashFromSizes -Sizes @($list | ForEach-Object { [long]$_.Length })
        Format    = $format
        FileCount = $list.Count
    }
}

# Single seam for every TheDiscDB HTTP call - tests replace this function with a mock.
# Short timeout: the lookup happens before the rip starts and must never hold it up.
function Invoke-DiscDbRequest {
    param([string]$Body)
    return Invoke-RestMethod -Uri $script:DiscDbEndpoint -Method Post -ContentType 'application/json' -Body $Body -TimeoutSec 8
}

# "0:21:35" -> 1295. $null for anything unparseable.
function ConvertFrom-DiscDbDuration {
    param([string]$Text)
    if ("$Text" -match '^\s*(?:(\d+):)?(\d{1,2}):(\d{2})\s*$') {
        $h = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
        return ($h * 3600) + ([int]$Matches[2] * 60) + [int]$Matches[3]
    }
    return $null
}

function ConvertTo-DiscDbInt {
    param($Value)
    if ("$Value" -match '^\s*(\d+)\s*$') { return [int]$Matches[1] }
    return $null
}

# Looks the content hash up. Returns the disc record, or $null when TheDiscDB has no
# disc with that hash. Throws on transport errors and GraphQL errors - the caller
# (Get-SeriesDiscDbLookup) turns those into a silent fallback.
function Find-DiscDbDisc {
    param([string]$ContentHash)

    # Filters at every level, so only the one matching disc (with its titles) comes
    # back - about 3 KB instead of every disc of a complete-series box set.
    $query = 'query($hash: String!) { mediaItems(where: { releases: { some: { discs: { some: { contentHash: { eq: $hash } } } } } }) { nodes { title type year releases(where: { discs: { some: { contentHash: { eq: $hash } } } }) { slug title discs(where: { contentHash: { eq: $hash } }) { index name format contentHash titles { index sourceFile duration size itemType season episode item { title type season episode } } } } } } }'
    $body = @{ query = $query; variables = @{ hash = $ContentHash } } | ConvertTo-Json -Depth 5 -Compress
    $response = Invoke-DiscDbRequest -Body $body

    if ($null -eq $response) { throw "empty response" }
    if ($response.errors) { throw "TheDiscDB error: $(@($response.errors)[0].message)" }

    $found = @()
    foreach ($media in @($response.data.mediaItems.nodes)) {
        if ($null -eq $media) { continue }
        foreach ($release in @($media.releases)) {
            if ($null -eq $release) { continue }
            foreach ($disc in @($release.discs)) {
                if ($disc -and "$($disc.contentHash)" -ieq $ContentHash) {
                    $found += [pscustomobject]@{ Media = $media; Release = $release; Disc = $disc }
                }
            }
        }
    }
    if ($found.Count -eq 0) { return $null }

    # The same pressing can sit in several releases (e.g. a season set and a complete
    # set); the title mapping is per disc, so the first is as good as any.
    $pick = $found[0]
    $titles = @(foreach ($t in @($pick.Disc.titles)) {
        if ($null -eq $t) { continue }
        $item = $t.item
        $itemType = if ($item -and $item.type) { "$($item.type)" } elseif ($item -and $t.itemType) { "$($t.itemType)" } else { $null }
        $seasonText = if ($item -and $item.season) { $item.season } else { $t.season }
        $episodeText = if ($item -and $item.episode) { $item.episode } else { $t.episode }
        [pscustomobject]@{
            Index       = [int]$t.index
            SourceFile  = $t.sourceFile
            DurationSec = ConvertFrom-DiscDbDuration $t.duration
            Size        = $t.size
            # $null = TheDiscDB lists the title but has not identified it (play-all,
            # menu loop, warning screen...). Such a title still takes part in lining
            # files up, but never decides what a file is.
            ItemType    = $itemType
            Season      = if ($item) { ConvertTo-DiscDbInt $seasonText } else { $null }
            Episode     = if ($item) { ConvertTo-DiscDbInt $episodeText } else { $null }
            Name        = if ($item) { "$($item.title)" } else { "" }
        }
    })

    $mediaText = "$($pick.Media.title)$(if ($pick.Media.year) { " ($($pick.Media.year))" })"
    $releaseText = if ($pick.Release.title) { " - $($pick.Release.title)" } else { "" }
    $discText = if ($pick.Disc.name) { ", $($pick.Disc.name)" } else { ", disc $($pick.Disc.index)" }
    return [pscustomobject]@{
        ContentHash  = $ContentHash
        MediaTitle   = $pick.Media.title
        MediaType    = $pick.Media.type
        Year         = $pick.Media.year
        ReleaseTitle = $pick.Release.title
        ReleaseSlug  = $pick.Release.slug
        DiscIndex    = [int]$pick.Disc.index
        DiscName     = $pick.Disc.name
        Format       = $pick.Disc.format
        Description  = "$mediaText$releaseText$discText"
        MatchCount   = $found.Count
        Titles       = $titles
    }
}

# Fail-soft wrapper used by both scripts. ALWAYS returns an object:
#   ContentHash - the disc's hash ("" if it could not be read), kept even when the
#                 lookup fails so a continue-rip.ps1 run can try again later
#   Disc        - the Find-DiscDbDisc record, or $null
#   Message     - one line describing the outcome, for the session log
# Pass -ContentHash when it is already known (continue-rip.ps1), otherwise -DriveRoot.
# Nothing here throws; every failure falls back to TMDb / median with one line shown.
function Get-SeriesDiscDbLookup {
    param([string]$DriveRoot = "", [string]$ContentHash = "", [switch]$Disabled, [string]$NoDriveReason = "")

    $result = [pscustomobject]@{ ContentHash = "$ContentHash".Trim().ToUpper(); Disc = $null; Message = "" }

    if ($Disabled) {
        $result.Message = "TheDiscDB: lookup off (-NoDiscDb) - episode/extra detection uses TMDb / title lengths"
    } else {
        try {
            if (-not $result.ContentHash) {
                if (-not $DriveRoot) {
                    $reason = if ($NoDriveReason) { $NoDriveReason } else { 'no disc hash available' }
                    $result.Message = "TheDiscDB: $reason - using TMDb / title lengths"
                } else {
                    $hashInfo = Get-DiscDbContentHash -DriveRoot $DriveRoot
                    if ($hashInfo) {
                        $result.ContentHash = $hashInfo.Hash
                    } else {
                        $result.Message = "TheDiscDB: no BDMV\STREAM or VIDEO_TS readable on $DriveRoot - using TMDb / title lengths"
                    }
                }
            } elseif ($result.ContentHash -notmatch '^[0-9A-F]{32}$') {
                $result.Message = "TheDiscDB: '$ContentHash' is not a valid disc hash (32 hex characters) - using TMDb / title lengths"
                $result.ContentHash = ""
            }

            if (-not $result.Message) {
                $disc = Find-DiscDbDisc -ContentHash $result.ContentHash
                if ($disc) {
                    $identified = @($disc.Titles | Where-Object { $_.ItemType }).Count
                    $result.Disc = $disc
                    $result.Message = "TheDiscDB: matched $($disc.Description) [$($result.ContentHash)] - $identified of $(@($disc.Titles).Count) title(s) identified"
                    Write-Host $result.Message -ForegroundColor Gray
                    return $result
                }
                $result.Message = "TheDiscDB: disc $($result.ContentHash) is not in the database - using TMDb / title lengths"
            }
        } catch {
            $result.Disc = $null
            $result.Message = "TheDiscDB lookup failed ($($_.Exception.Message)) - using TMDb / title lengths"
        }
    }
    Write-Host $result.Message -ForegroundColor DarkGray
    return $result
}

# Lowest episode number TheDiscDB lists on this disc for -Season (any season when
# -Season is 0). Used as the default answer at the Disc 2+ start-episode prompt.
function Get-DiscDbFirstEpisode {
    param([object]$DiscDbDisc, [int]$Season = 0)
    if (-not $DiscDbDisc) { return $null }
    $eps = @($DiscDbDisc.Titles | Where-Object {
        $_.ItemType -eq 'Episode' -and $null -ne $_.Episode -and $_.Episode -ge 1 -and
        ($Season -le 0 -or $null -eq $_.Season -or $_.Season -eq $Season)
    } | ForEach-Object { $_.Episode })
    if ($eps.Count -eq 0) { return $null }
    return [int](($eps | Measure-Object -Minimum).Minimum)
}

# Lines the ripped files up with TheDiscDB's titles and returns
# @{ <file name> = <TheDiscDB title> } for each file that lines up.
#
# TheDiscDB's title index is MakeMKV's own title number, but the numbers in our file
# names (title_t03) can drift from it: MakeMKV's minimum-length setting drops short
# titles (TheDiscDB lists even 9-second clips), Step 2 skips the play-all composite,
# and MakeMKV versions differ. Durations alone are not enough either - episodes on one
# disc are often within a second or two of each other. So this is an order-preserving
# alignment (dynamic programming, like a diff): files and TheDiscDB titles are both in
# MakeMKV order, a file may pair with a title only if their durations agree within
# max(-ToleranceMinSec, -TolerancePct), and the alignment with the most pairs wins;
# ties go to the closest durations, then to the smallest title-number drift.
function Get-DiscDbTitleMatches {
    param(
        [object[]]$Titles,
        [object]$DiscDbDisc,
        [double]$TolerancePct = 0.01,
        [double]$ToleranceMinSec = 5
    )
    $map = @{}
    if (-not $DiscDbDisc -or -not $Titles -or $Titles.Count -eq 0) { return $map }
    $dd = @($DiscDbDisc.Titles | Where-Object { $null -ne $_ } | Sort-Object Index)
    if ($dd.Count -eq 0) { return $map }

    $files = @($Titles)
    $n = $files.Count
    $m = $dd.Count
    $fileIndex = @(for ($i = 0; $i -lt $n; $i++) {
        if ("$($files[$i].Name)" -match '_t(\d+)\.[^.]+$') { [int]$Matches[1] } else { $i }
    })

    $score = New-Object 'double[,]' ($n + 1), ($m + 1)
    $move = New-Object 'int[,]' ($n + 1), ($m + 1)   # 1 = skip file, 2 = skip title, 3 = pair
    for ($i = 1; $i -le $n; $i++) { $move[$i, 0] = 1 }
    for ($j = 1; $j -le $m; $j++) { $move[0, $j] = 2 }

    for ($i = 1; $i -le $n; $i++) {
        $fd = $files[$i - 1].DurationSec
        for ($j = 1; $j -le $m; $j++) {
            $best = $score[($i - 1), $j]; $mv = 1
            if ($score[$i, ($j - 1)] -gt $best) { $best = $score[$i, ($j - 1)]; $mv = 2 }
            $td = $dd[$j - 1].DurationSec
            if ($null -ne $fd -and [double]$fd -gt 0 -and $null -ne $td -and $td -gt 0) {
                $diff = [math]::Abs([double]$fd - [double]$td)
                if ($diff -le [math]::Max($ToleranceMinSec, $td * $TolerancePct)) {
                    $w = 10000 - ($diff * 10) - [math]::Abs($dd[$j - 1].Index - $fileIndex[$i - 1])
                    if ($score[($i - 1), ($j - 1)] + $w -gt $best) { $best = $score[($i - 1), ($j - 1)] + $w; $mv = 3 }
                }
            }
            $score[$i, $j] = $best
            $move[$i, $j] = $mv
        }
    }

    $i = $n; $j = $m
    while ($i -gt 0 -and $j -gt 0) {
        switch ($move[$i, $j]) {
            3 { $map[$files[$i - 1].Name] = $dd[$j - 1]; $i--; $j-- }
            2 { $j-- }
            default { $i-- }
        }
    }
    return $map
}
