# SeriesRetroRename.ps1 - folder discovery and orchestration for rename-series.ps1, the
# standalone series rename utility for episodes that are ALREADY ripped on disk.
#
# Dot-sourced by rename-series.ps1 AFTER SeriesEpisodes.ps1. It adds no naming, no
# classification and no manifest logic of its own: every folder goes through the same
# Invoke-SeriesEpisodeRename a live -Series rip uses (episode vs extra criteria, TMDb
# runtimes, <Title>-S##-E## / extras\<Title>-S##-Extra## names, the confirmation table,
# rename-manifest.csv written BEFORE any file moves, undo-rename.ps1 copied next to it,
# never overwriting). This file only works out WHICH folders to run it on, in what order,
# and where each folder's episode numbering starts.
#
# Layout it understands (the one rip-disc.ps1 produces):
#   <root>\<Title>\Season N\DiscN\*.mkv|mp4
#   <root>\<Title>\DiscN\...            (a series ripped without -Season: S01 fallback)
# and accepts the root at any level: the series folder, one Season folder, or one DiscN.

# Finds the folders to process. Only Season N / DiscN folders become units, so a folder
# named "extras" (already-moved extras) is never treated as one.
function Get-SeriesRenameUnits {
    param(
        [string]$Root,
        # Overrides the title taken from the folder name (the series folder, or the folder
        # above Season N / DiscN).
        [string]$Title = "",
        # Season to use when the path does not say (no "Season N" folder). 0 = fall back to S01.
        [int]$Season = 0
    )

    $Root = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('\')
    $units = New-Object System.Collections.Generic.List[object]

    $addDiscUnits = {
        param([string]$ParentDir, [string]$UnitTitle, [int]$UnitSeason)
        $discs = @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^Disc(\d+)$' } |
            Sort-Object { [int]($_.Name -replace '\D', '') })
        foreach ($d in $discs) {
            $units.Add([pscustomobject]@{
                Directory = $d.FullName
                Title     = $UnitTitle
                Season    = $UnitSeason
                Disc      = [int]($d.Name -replace '\D', '')
                SeasonDir = $ParentDir
            })
        }
        return $discs.Count
    }

    $leaf = Split-Path $Root -Leaf
    $parent = Split-Path $Root -Parent
    $parentLeaf = if ($parent) { Split-Path $parent -Leaf } else { "" }

    if ($leaf -match '^Disc(\d+)$') {
        $discNumber = [int]$Matches[1]
        # A single Disc folder: season from a "Season N" parent, title from the folder above that.
        if ($parentLeaf -match '^Season\s+(\d+)$') {
            $s = [int]$Matches[1]
            $t = if ($Title) { $Title } else { Split-Path (Split-Path $parent -Parent) -Leaf }
        } else {
            $s = $Season
            $t = if ($Title) { $Title } else { $parentLeaf }
        }
        $units.Add([pscustomobject]@{ Directory = $Root; Title = $t; Season = $s; Disc = $discNumber; SeasonDir = $parent })
    } elseif ($leaf -match '^Season\s+(\d+)$') {
        $s = [int]$Matches[1]
        $t = if ($Title) { $Title } else { $parentLeaf }
        $found = & $addDiscUnits $Root $t $s
        if ($found -eq 0) {
            # No Disc folders: the Season folder itself holds the files.
            $units.Add([pscustomobject]@{ Directory = $Root; Title = $t; Season = $s; Disc = 0; SeasonDir = $Root })
        }
    } else {
        # A series folder: Season N children (and/or Disc children when there is no Season folder).
        $t = if ($Title) { $Title } else { $leaf }
        $seasonDirs = @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^Season\s+(\d+)$' } |
            Sort-Object { [int]($_.Name -replace '\D', '') })
        foreach ($sd in $seasonDirs) {
            $s = [int]($sd.Name -replace '\D', '')
            $found = & $addDiscUnits $sd.FullName $t $s
            if ($found -eq 0) {
                $units.Add([pscustomobject]@{ Directory = $sd.FullName; Title = $t; Season = $s; Disc = 0; SeasonDir = $sd.FullName })
            }
        }
        [void](& $addDiscUnits $Root $t $Season)
    }

    $out = @()
    foreach ($x in $units) { $out += $x }
    return $out
}

# Highest episode number among files in $Directory already in final <Title>-S##-E## shape
# (0 when none), so numbering carries on after episodes renamed by an earlier run.
function Get-SeriesFolderMaxEpisode {
    param([string]$Directory, [string]$Title, [int]$Season)
    $pattern = Get-SeriesNamePattern -Title $Title -Season $Season
    $max = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $Directory -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match $pattern -and $Matches['ep']) {
            $n = [int]$Matches['ep']
            if ($n -gt $max) { $max = $n }
        }
    }
    return $max
}

# Runs every unit under -Root through Invoke-SeriesEpisodeRename, in season then disc
# order. Dry run (no -Apply): shows each folder's planned names and changes nothing.
# Episode numbers continue across the discs of a season - from the earlier discs' renamed
# files / manifests when they are already renamed, and from the plan itself when they are
# only planned (a dry run, or an apply that was declined for an earlier disc).
function Invoke-SeriesRetroRename {
    param(
        [Parameter(Mandatory)][string]$Root,
        [string]$Title = "",
        [int]$Season = 0,
        # Starting episode for the first folder of each season (default: 1).
        [int]$StartEpisode = 0,
        [switch]$Apply,
        # With -Apply: accept every folder's confirmation table without asking.
        [switch]$Yes,
        [switch]$NoTmdb,
        [string]$HandBrakePath = "",
        [string]$UndoScriptSource = "",
        [scriptblock]$ReadInput = { param($p) Read-Host $p },
        [scriptblock]$GetDuration = $null,
        # Injectable TMDb season lookup (tests); default is the real one.
        [scriptblock]$GetTmdbSeason = $null
    )

    $units = @(Get-SeriesRenameUnits -Root $Root -Title $Title -Season $Season)
    if ($units.Count -eq 0) {
        Write-Host "No Season / DiscN folders found under $Root" -ForegroundColor Yellow
        return [pscustomobject]@{ Units = 0; Planned = 0; Renamed = 0; Skipped = 0; Declined = 0; Results = @() }
    }

    $mode = if ($Apply) { "APPLY" } else { "DRY RUN - nothing will be renamed (add -Apply to rename)" }
    Write-Host "`n$mode" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Cyan' })
    Write-Host "$($units.Count) folder(s) to process under $Root" -ForegroundColor Gray

    $nextBySeason = @{}
    $tmdbBySeason = @{}
    $results = @()
    foreach ($u in $units) {
        $seasonText = if ($u.Season -gt 0) { "Season $($u.Season)" } else { "no season (S01)" }
        Write-Host "`n=== $($u.Title) - $seasonText - $(Split-Path $u.Directory -Leaf) ===" -ForegroundColor Cyan
        Write-Host "  $($u.Directory)" -ForegroundColor Gray

        $key = "$($u.Season)"
        $start = 0
        $fromWhere = ""
        if ($nextBySeason.ContainsKey($key)) { $start = $nextBySeason[$key]; $fromWhere = "continuing after the previous folder" }
        if ($u.Disc -ge 2) {
            $suggested = Get-SuggestedStartEpisode -SeasonDir $u.SeasonDir -Disc $u.Disc
            if ($suggested -and [int]$suggested -gt $start) { $start = [int]$suggested; $fromWhere = "after the episodes on earlier discs" }
        }
        if ($start -eq 0) {
            if ($StartEpisode -ge 1) {
                $start = $StartEpisode
                $fromWhere = "-StartEpisode"
            } else {
                $start = 1
                if ($u.Disc -ge 2) {
                    Write-Host "  WARNING: no episodes found on earlier discs - numbering starts at E01. Pass -StartEpisode N if this disc continues a season." -ForegroundColor Yellow
                }
            }
        }
        if ($fromWhere) { Write-Host ("  Starting at E{0:D2} ({1})" -f $start, $fromWhere) -ForegroundColor Gray }

        $tmdb = $null
        if (-not $NoTmdb) {
            if (-not $tmdbBySeason.ContainsKey($key)) {
                $tmdbBySeason[$key] = if ($GetTmdbSeason) { & $GetTmdbSeason $u.Title $u.Season } else { Get-SeriesTmdbSeason -Title $u.Title -Season $u.Season }
            }
            $tmdb = $tmdbBySeason[$key]
        }

        $renameArgs = @{
            Directory        = $u.Directory
            Title            = $u.Title
            Season           = $u.Season
            StartEpisode     = $start
            TmdbSeason       = $tmdb
            HandBrakePath    = $HandBrakePath
            UndoScriptSource = $UndoScriptSource
            ReadInput        = $ReadInput
            AutoAccept       = [bool]$Yes
            DryRun           = (-not $Apply)
        }
        if ($GetDuration) { $renameArgs.GetDuration = $GetDuration }
        $r = Invoke-SeriesEpisodeRename @renameArgs
        $results += [pscustomobject]@{ Unit = $u; Result = $r }

        # Carry numbering forward: planned (not skipped) episodes plus anything already renamed.
        $planMax = 0
        foreach ($p in @($r.Plan | Where-Object { $_.Kind -eq 'Episode' -and -not $_.Skip })) {
            $m = [regex]::Match($p.NewName, '-E(\d+)\.[^.\\]+$')
            if ($m.Success -and [int]$m.Groups[1].Value -gt $planMax) { $planMax = [int]$m.Groups[1].Value }
        }
        $folderMax = Get-SeriesFolderMaxEpisode -Directory $u.Directory -Title $u.Title -Season $u.Season
        $next = [math]::Max($planMax, $folderMax) + 1
        if ($next -gt 1 -and (-not $nextBySeason.ContainsKey($key) -or $next -gt $nextBySeason[$key])) { $nextBySeason[$key] = $next }
    }

    $planned = 0; $renamed = 0; $skipped = 0; $declined = 0
    foreach ($x in $results) {
        $planned += @($x.Result.Plan | Where-Object { -not $_.Skip }).Count
        $renamed += [int]$x.Result.Renamed
        $skipped += [int]$x.Result.Skipped
        if ($x.Result.DryRun) { $skipped += @($x.Result.Plan | Where-Object { $_.Skip }).Count }
        if ($x.Result.Declined) { $declined++ }
    }
    Write-Host ""
    if ($Apply) {
        $tail = ""
        if ($skipped) { $tail += ", skipped $skipped" }
        if ($declined) { $tail += ", $declined folder(s) declined" }
        Write-Host "Done: renamed $renamed file(s) in $($units.Count) folder(s)$tail." -ForegroundColor Green
    } else {
        $tail = if ($skipped) { " ($skipped would be skipped - target exists)" } else { "" }
        Write-Host "Dry run complete: $planned file(s) would be renamed in $($units.Count) folder(s)$tail. Nothing was changed." -ForegroundColor Cyan
        Write-Host "Re-run with -Apply to rename; each folder gets rename-manifest.csv + undo-rename.ps1 first." -ForegroundColor Gray
    }
    return [pscustomobject]@{ Units = $units.Count; Planned = $planned; Renamed = $renamed; Skipped = $skipped; Declined = $declined; Results = $results }
}
