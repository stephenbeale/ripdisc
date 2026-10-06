# SeriesRetroRename.ps1 - folder discovery and orchestration for rename-series.ps1, the
# standalone series rename utility for episodes that are ALREADY ripped on disk.
#
# Dot-sourced by rename-series.ps1 AFTER SeriesEpisodes.ps1. It adds no naming, no
# classification and no manifest logic of its own: every folder goes through the same
# Invoke-SeriesEpisodeRename a live -Series rip uses (episode vs extra criteria, TMDb
# runtimes, <Title>-S##-E## / Specials\<Title>-S##-D#-Extra## names, the confirmation table,
# rename-manifest.csv written BEFORE any file moves, undo-rename.ps1 copied next to it,
# never overwriting). This file only works out WHICH folders to run it on, in what order,
# and where each folder's episode numbering starts.
#
# Layout it understands (the one rip-disc.ps1 produces):
#   <root>\<Title>\Season N\DiscN\*.mkv|mp4
#   <root>\<Title>\DiscN\...            (a series ripped without -Season: S01 fallback)
# and accepts the root at any level: the series folder, one Season folder, or one DiscN.

# Season number from a folder name, or $null. Accepts "Season 2", "Series 1", and a
# trailing form after any prefix: "Joking Apart-Series 1", "Joking Apart Series 1",
# "Show - Season 3" (case-insensitive, "Series1" too).
function Get-SeasonFolderNumber {
    param([string]$Name)
    if ($Name -match '(?i)(?:^|[\s\-_.])(?:series|season)\s*(\d+)$') { return [int]$Matches[1] }
    return $null
}

# Disc number from a Disc folder name ("Disc1" or "Disc 1"), or $null.
function Get-DiscFolderNumber {
    param([string]$Name)
    if ($Name -match '(?i)^disc\s*(\d+)$') { return [int]$Matches[1] }
    return $null
}

# Disc number from a "Disc N" token inside a FILE name ("Joking Apart-Series 2 Disc 1-C1_t00.mp4"), or $null.
function Get-DiscNumberFromFileName {
    param([string]$Name)
    if ($Name -match '(?i)\bdisc\s*(\d+)(?!\d)') { return [int]$Matches[1] }
    return $null
}

# Finds the folders to process. Only season / Disc folders become units, so a folder
# named "Specials" or "extras" (already-moved extras) is never treated as one and its
# files are never renamed here.
#
# A season folder with no Disc subfolders holds its files directly. When their names carry
# a "Disc N" token (older rips put several discs in one folder), each disc becomes its own
# unit - same folder, a FileFilter picking that disc's files - so numbering continues
# across them. Nothing is moved into new Disc folders.
function Get-SeriesRenameUnits {
    param(
        [string]$Root,
        # Overrides the title taken from the folder name (the series folder, or the folder
        # above the season / Disc folder).
        [string]$Title = "",
        # Season to use when the path does not say (no season folder). 0 = fall back to S01.
        [int]$Season = 0
    )

    $Root = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('')
    $units = New-Object System.Collections.Generic.List[object]

    $addDiscUnits = {
        param([string]$ParentDir, [string]$UnitTitle, [int]$UnitSeason)
        $discs = @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $null -ne (Get-DiscFolderNumber $_.Name) } |
            Sort-Object { Get-DiscFolderNumber $_.Name })
        foreach ($d in $discs) {
            $units.Add([pscustomobject]@{
                Directory = $d.FullName
                Title     = $UnitTitle
                Season    = $UnitSeason
                Disc      = (Get-DiscFolderNumber $d.Name)
                SeasonDir = $ParentDir
                FileFilter = $null
            })
        }
        return $discs.Count
    }

    # A season folder with no Disc subfolders: one unit, or one per "Disc N" file-name token.
    $addFlatSeasonUnits = {
        param([string]$Dir, [string]$UnitTitle, [int]$UnitSeason)
        $pattern = Get-SeriesNamePattern -Title $UnitTitle -Season $UnitSeason
        $videos = @(Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^\.(mp4|mkv)$' -and $_.Name -notmatch $pattern })
        $tokens = @($videos | ForEach-Object { Get-DiscNumberFromFileName $_.Name } | Where-Object { $null -ne $_ } | Sort-Object -Unique)
        if ($tokens.Count -eq 0) {
            $units.Add([pscustomobject]@{ Directory = $Dir; Title = $UnitTitle; Season = $UnitSeason; Disc = 0; SeasonDir = $Dir; FileFilter = $null })
            return
        }
        # Files with no token (if any) come first as their own unit, then each disc in order.
        if (@($videos | Where-Object { $null -eq (Get-DiscNumberFromFileName $_.Name) }).Count -gt 0) {
            $units.Add([pscustomobject]@{
                Directory = $Dir; Title = $UnitTitle; Season = $UnitSeason; Disc = 0; SeasonDir = $Dir
                FileFilter = { param($n) $null -eq (Get-DiscNumberFromFileName $n) }
            })
        }
        # GetNewClosure() binds the filter to a new dynamic module, which cannot see functions
        # dot-sourced into rename-series.ps1's script scope - so capture the function itself.
        $discOf = ${function:Get-DiscNumberFromFileName}
        foreach ($t in $tokens) {
            $wanted = [int]$t
            $units.Add([pscustomobject]@{
                Directory = $Dir; Title = $UnitTitle; Season = $UnitSeason; Disc = $wanted; SeasonDir = $Dir
                FileFilter = { param($n) (& $discOf $n) -eq $wanted }.GetNewClosure()
            })
        }
    }

    $leaf = Split-Path $Root -Leaf
    $parent = Split-Path $Root -Parent
    $parentLeaf = if ($parent) { Split-Path $parent -Leaf } else { "" }
    $leafSeason = Get-SeasonFolderNumber $leaf
    $leafDisc = Get-DiscFolderNumber $leaf

    if ($null -ne $leafDisc) {
        # A single Disc folder: season from a season-folder parent, title from the folder above that.
        $parentSeason = Get-SeasonFolderNumber $parentLeaf
        if ($null -ne $parentSeason) {
            $s = $parentSeason
            $t = if ($Title) { $Title } else { Split-Path (Split-Path $parent -Parent) -Leaf }
        } else {
            $s = $Season
            $t = if ($Title) { $Title } else { $parentLeaf }
        }
        $units.Add([pscustomobject]@{ Directory = $Root; Title = $t; Season = $s; Disc = $leafDisc; SeasonDir = $parent; FileFilter = $null })
    } elseif ($null -ne $leafSeason) {
        $t = if ($Title) { $Title } else { $parentLeaf }
        $found = & $addDiscUnits $Root $t $leafSeason
        if ($found -eq 0) { & $addFlatSeasonUnits $Root $t $leafSeason }
    } else {
        # A series folder: season children (and/or Disc children when there is no season folder).
        $t = if ($Title) { $Title } else { $leaf }
        $seasonDirs = @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $null -ne (Get-SeasonFolderNumber $_.Name) } |
            Sort-Object { Get-SeasonFolderNumber $_.Name })
        foreach ($sd in $seasonDirs) {
            $s = Get-SeasonFolderNumber $sd.Name
            $found = & $addDiscUnits $sd.FullName $t $s
            if ($found -eq 0) { & $addFlatSeasonUnits $sd.FullName $t $s }
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
    $reservedExtrasByDir = @{}
    $results = @()
    foreach ($u in $units) {
        $seasonText = if ($u.Season -gt 0) { "Season $($u.Season)" } else { "no season (S01)" }
        $discText = if ($u.FileFilter) { if ($u.Disc -gt 0) { " - files for Disc $($u.Disc)" } else { " - files with no disc number" } } else { "" }
        Write-Host "`n=== $($u.Title) - $seasonText - $(Split-Path $u.Directory -Leaf)$discText ===" -ForegroundColor Cyan
        Write-Host "  $($u.Directory)" -ForegroundColor Gray

        $key = "$($u.Season)"
        $start = 0
        $fromWhere = ""
        if ($nextBySeason.ContainsKey($key)) { $start = $nextBySeason[$key]; $fromWhere = "continuing after the previous folder" }
        if ($u.Disc -ge 2) {
            $suggested = Get-SuggestedStartEpisode -SeasonDir $u.SeasonDir -Disc $u.Disc
            if ($suggested -and [int]$suggested -gt $start) { $start = [int]$suggested; $fromWhere = "after the episodes on earlier discs" }
        }
        if ($u.FileFilter -and $u.Disc -ge 2) {
            # Several discs share this folder: carry on after any already-renamed episodes in it.
            $inFolder = Get-SeriesFolderMaxEpisode -Directory $u.Directory -Title $u.Title -Season $u.Season
            if ($inFolder + 1 -gt $start) { $start = $inFolder + 1; $fromWhere = "after the episodes already renamed in this folder" }
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
        if ($u.FileFilter) { $renameArgs.FileFilter = $u.FileFilter }
        $dirKey = $u.Directory.ToLowerInvariant()
        if ($reservedExtrasByDir.ContainsKey($dirKey)) { $renameArgs.ReservedExtras = [int[]]@($reservedExtrasByDir[$dirKey]) }
        $r = Invoke-SeriesEpisodeRename @renameArgs
        $results += [pscustomobject]@{ Unit = $u; Result = $r }

        # Carry numbering forward: planned (not skipped) episodes plus anything already renamed.
        $planMax = 0
        foreach ($p in @($r.Plan | Where-Object { $_.Kind -eq 'Episode' -and -not $_.Skip })) {
            $m = [regex]::Match($p.NewName, '-E(\d+)\.[^.\\]+$')
            if ($m.Success -and [int]$m.Groups[1].Value -gt $planMax) { $planMax = [int]$m.Groups[1].Value }
        }
        foreach ($p in @($r.Plan | Where-Object { $_.Kind -eq 'Extra' -and -not $_.Skip })) {
            $em = [regex]::Match($p.NewName, '-Extra(\d+)')
            if ($em.Success) {
                if (-not $reservedExtrasByDir.ContainsKey($dirKey)) { $reservedExtrasByDir[$dirKey] = @() }
                $reservedExtrasByDir[$dirKey] += [int]$em.Groups[1].Value
            }
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
