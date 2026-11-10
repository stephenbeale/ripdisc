#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the mediaServer setting ("jellyfin" / "plex"): how Load-Config.ps1 reads it,
    the movie main-file and extras-folder names, plain -Series episode / special / extra
    names and folders, genre-series episode names, undo of a Plex-style rename, and that
    setup.ps1 and the sample config carry the setting.

.DESCRIPTION
    Lifts the real functions and Step 3 blocks out of the scripts with the AST parser /
    script text, so nothing here is a reimplementation. Filesystem cases use a temp
    directory removed afterwards. No network, disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-MediaServerNaming.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$configPath = Join-Path $repoRoot 'Load-Config.ps1'
$sharedPath = Join-Path $repoRoot 'SeriesEpisodes.ps1'
$retroPath = Join-Path $repoRoot 'SeriesRetroRename.ps1'
$undoPath = Join-Path $repoRoot 'undo-rename.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'
$setupPath = Join-Path $repoRoot 'setup.ps1'
$samplePath = Join-Path $repoRoot 'ripdisc-config.sample.json'

foreach ($p in @($configPath, $sharedPath, $retroPath, $undoPath, $ripDiscPath, $continuePath, $setupPath, $samplePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }
    $fn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

foreach ($name in @('Get-FeatureNamePattern', 'Get-MovieExtrasFolderName', 'Get-AnyMainFilePattern', 'Get-ExistingFeatureFile')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
}
$script:AnyFeaturePattern = '-Feature(-BluRay)?\.[^.]+$'
. $sharedPath
. ([scriptblock]::Create((Import-FunctionFromScript $ripDiscPath 'Get-EpisodeFileName')))

$script:LogLines = New-Object System.Collections.Generic.List[string]
function Write-Log { param([string]$Message) $script:LogLines.Add($Message) }

function New-InputQueue {
    param([object[]]$Answers)
    $queue = New-Object System.Collections.Queue
    foreach ($a in $Answers) { $queue.Enqueue($a) }
    return { param($prompt) if ($queue.Count -gt 0) { $queue.Dequeue() } else { $null } }.GetNewClosure()
}

function Get-RelNames {
    param([string]$Dir)
    $root = (Resolve-Path -LiteralPath $Dir).ProviderPath.TrimEnd('\') + '\'
    (Get-ChildItem -LiteralPath $Dir -File -Recurse | Where-Object { $_.Extension -in '.mp4', '.mkv' } |
        ForEach-Object { $_.FullName.Substring($root.Length) } | Sort-Object) -join ','
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

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-mediaserver-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
$savedEnvKey = $env:TMDB_API_KEY

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nReading the setting (Load-Config.ps1)" -ForegroundColor Cyan

    $cfgLine = @(Get-Content -LiteralPath $configPath | Where-Object { $_ -match '^\$script:Config_MediaServer = ' })
    Assert-Equal 1 $cfgLine.Count 'Load-Config.ps1 sets $script:Config_MediaServer in one line'
    foreach ($case in @(@{ V = 'plex'; E = 'plex' }, @{ V = ' Plex '; E = 'plex' }, @{ V = 'jellyfin'; E = 'jellyfin' },
                        @{ V = 'JELLYFIN'; E = 'jellyfin' }, @{ V = 'emby'; E = 'jellyfin' }, @{ V = ''; E = 'jellyfin' })) {
        $cfg = [pscustomobject]@{ mediaServer = $case.V }
        Invoke-Expression $cfgLine[0]
        Assert-Equal $case.E $script:Config_MediaServer "mediaServer `"$($case.V)`" -> $($case.E)"
    }
    $cfg = [pscustomobject]@{ tempRoot = 'C:\Video' }
    Invoke-Expression $cfgLine[0]
    Assert-Equal 'jellyfin' $script:Config_MediaServer 'an older config without mediaServer stays on Jellyfin naming'
    $cfg = $null
    Invoke-Expression $cfgLine[0]
    Assert-Equal 'jellyfin' $script:Config_MediaServer 'no config file at all -> Jellyfin naming'

    # ---------------------------------------------------------------------------
    Write-Host "`nMovies: main file and extras folder" -ForegroundColor Cyan

    $script:Config_MediaServer = 'jellyfin'
    Assert-Equal 'extras' (Get-MovieExtrasFolderName) 'Jellyfin extras folder is still "extras"'
    Assert-Equal '-Feature\.[^.]+$' (Get-FeatureNamePattern -FolderName 'Heat') 'Jellyfin main-file pattern is unchanged when a folder name is passed'
    Assert-Equal '-Part 2\.[^.]+$' (Get-FeatureNamePattern -Part 2 -FolderName 'Heat') 'Jellyfin part pattern is unchanged'

    $script:Config_MediaServer = 'plex'
    Assert-Equal 'Other' (Get-MovieExtrasFolderName) 'Plex extras go in "Other" (Plex does not read "extras")'
    Assert-Equal 'extras' (Get-MovieExtrasFolderName -MediaServer 'jellyfin') '-MediaServer overrides the setting'
    $dvd = Get-FeatureNamePattern -FolderName 'Heat (1995)'
    Assert-True ('Heat (1995).mp4' -match $dvd) 'Plex DVD main file is named after its folder: Heat (1995).mp4'
    Assert-True ('Heat (1995) - BluRay.mp4' -notmatch $dvd) '... and the Blu-ray version is not taken for it'
    Assert-True ('Heat (1995)-Feature.mp4' -notmatch $dvd) '... nor a Jellyfin-style Feature'
    Assert-True ('Heat (1995)-title_t00.mp4' -notmatch $dvd) '... nor an ordinary encode'
    Assert-True ('Heat (1995)x.mp4' -notmatch $dvd) 'brackets and dots in the folder name are matched literally'
    $bd = Get-FeatureNamePattern -Bluray -FolderName 'Heat (1995)'
    Assert-True ('Heat (1995) - BluRay.mkv' -match $bd -and 'Heat (1995).mkv' -notmatch $bd) 'Plex Blu-ray main file: Heat (1995) - BluRay.mkv, beside the DVD version'
    $pt = Get-FeatureNamePattern -Part 2 -FolderName 'Cream'
    Assert-True ('Cream - pt2.mp4' -match $pt -and 'Cream - pt1.mp4' -notmatch $pt -and 'Cream - pt12.mp4' -notmatch $pt) 'Plex documentary part: Cream - pt2.mp4 (Plex''s multi-part token)'
    $threw = $false
    try { Get-FeatureNamePattern | Out-Null } catch { $threw = $true }
    Assert-True $threw 'Plex pattern without -FolderName throws rather than matching anything'

    $any = Get-AnyMainFilePattern -FolderName 'Heat'
    foreach ($n in @('Heat.mp4', 'Heat - BluRay.mkv', 'Heat - pt3.mp4', 'Heat-Feature.mp4', 'Heat-Feature-BluRay.mkv', 'Heat-Part 2.mp4')) {
        Assert-True ($n -match $any) "any-main-file pattern covers $n (either server's form)"
    }
    foreach ($n in @('Heat-title_t00.mp4', 'Heat-Special Features-title_t01.mp4', 'Heat 2.mp4', 'title_t00.mp4')) {
        Assert-True ($n -notmatch $any) "$n is still sorted (moved to extras / picked as largest)"
    }
    Assert-True ('Heat-Feature.mp4' -match (Get-AnyMainFilePattern) -and 'Heat.mp4' -notmatch (Get-AnyMainFilePattern)) 'no folder name: only the suffix forms'

    $heat = Join-Path $tempRoot 'DVDs\Heat'
    New-Item -ItemType Directory -Path $heat -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $heat 'Heat.mp4') -Value 'x'
    Assert-Equal 'Heat.mp4' (Get-ExistingFeatureFile -FolderPath $heat -Pattern (Get-FeatureNamePattern -FolderName 'Heat')) 'clash check finds an existing Plex main file'
    Assert-Equal $null (Get-ExistingFeatureFile -FolderPath $heat -Pattern (Get-FeatureNamePattern -Bluray -FolderName 'Heat')) 'a Blu-ray run does not see the Plex DVD file as its own'

    # The real Step 3 blocks from both scripts, in Plex mode
    foreach ($path in @($ripDiscPath, $continuePath)) {
        $leaf = Split-Path $path -Leaf
        $text = Get-Content -LiteralPath $path -Raw
        $folderLine = @(Get-Content -LiteralPath $path | Where-Object { $_ -match '^\$mainFolderName = ' })
        Assert-Equal 1 $folderLine.Count "$leaf works out the film's folder name once"
        $Series = $false
        $Extras = $true; $finalOutputDir = 'F:\DVDs\Heat (1995)\Other'
        Invoke-Expression $folderLine[0]
        Assert-Equal 'Heat (1995)' $mainFolderName "$leaf -Extras: the film folder is the extras folder's parent"
        $Extras = $false; $finalOutputDir = 'F:\DVDs\Heat (1995)'
        Invoke-Expression $folderLine[0]
        Assert-Equal 'Heat (1995)' $mainFolderName "$leaf Disc 1: the film folder is the output folder"

        $s = $text.IndexOf('# Naming of the main file')
        $e = $text.IndexOf('$extrasDir = Join-Path $finalOutputDir (Get-MovieExtrasFolderName)', $s)
        Assert-True ($s -ge 0 -and $e -gt $s) "$leaf has the main-file naming block and the extras folder line"
        $patBlock = [scriptblock]::Create($text.Substring($s, $e - $s))
        $mainFolderName = 'Cream'
        foreach ($case in @(@{ BD = $false; Part = $false; D = 1; Suffix = ''; Own = 'Cream.mp4'; Other = 'Cream - BluRay.mp4' },
                            @{ BD = $true; Part = $false; D = 1; Suffix = ' - BluRay'; Own = 'Cream - BluRay.mp4'; Other = 'Cream.mp4' },
                            @{ BD = $false; Part = $true; D = 2; Suffix = ' - pt2'; Own = 'Cream - pt2.mp4'; Other = 'Cream - pt1.mp4' })) {
            $Bluray = $case.BD; $Disc = $case.D; $script:DocumentaryPart = $case.Part
            . $patBlock
            $label = "$leaf Plex $(if ($case.Part) { 'part' } elseif ($case.BD) { 'Blu-ray' } else { 'DVD' })"
            Assert-Equal $case.Own ($mainFolderName + $featureSuffix + '.mp4') "$label`: the largest file is renamed to $($case.Own)"
            Assert-Equal $case.Suffix $featureSuffix "$label`: suffix"
            Assert-True ($case.Own -match $featurePattern -and $case.Other -notmatch $featurePattern) "$label`: its own main file counts as done, $($case.Other) does not"
            Assert-True ($case.Other -match $featureExcludePattern -and 'Cream-Feature.mp4' -match $featureExcludePattern) "$label`: other main files (incl. a Jellyfin-era Feature) are never swept into extras"
            Assert-True ('Cream-title_t03.mp4' -notmatch $featureExcludePattern) "$label`: ordinary encodes are still sorted"
        }
        $script:Config_MediaServer = 'jellyfin'
        $Bluray = $false; $Disc = 1; $script:DocumentaryPart = $false
        . $patBlock
        Assert-Equal '-Feature' $featureSuffix "$leaf Jellyfin: still -Feature"
        $script:Config_MediaServer = 'plex'

        Assert-True ($text.Contains('$finalOutputDir = Join-Path $finalOutputDir (Get-MovieExtrasFolderName)')) "$leaf -Extras encodes into the server's extras folder"
        Assert-True (-not $text.Contains('md extras') -and -not $text.Contains('Test-Path "extras"')) "$leaf has no hard-coded extras folder left in Step 3"
        Assert-True (-not $text.Contains("-notmatch '-Feature(-BluRay)?\.[^.]+$'")) "$leaf Disc 2+ keeps every main-file form out of extras"
    }
    $ripText = Get-Content -LiteralPath $ripDiscPath -Raw
    Assert-True ($ripText.Contains('-Pattern (Get-FeatureNamePattern -Bluray:$Bluray -Part $clashPart -FolderName $altName)')) 'rip-disc.ps1 pre-rip clash check builds the fallback folder''s own pattern'
    Assert-True ($ripText.Contains('Join-Path $outputFolderPreview (Get-MovieExtrasFolderName)')) 'rip-disc.ps1 Ready-to-rip preview shows the server''s extras folder'

    # ---------------------------------------------------------------------------
    Write-Host "`nSeries: episode, special and extra names" -ForegroundColor Cyan

    $script:Config_MediaServer = 'jellyfin'
    Assert-Equal 'Fargo-S01-E03.mp4' (Get-SeriesEpisodeFileName -Title 'Fargo' -Season 1 -Episode 3 -Extension '.mp4') 'Jellyfin episode name unchanged'
    Assert-Equal 'Fargo-S00-E02.mkv' (Get-SeriesSpecialFileName -Title 'Fargo' -Special 2 -Extension '.mkv') 'Jellyfin special name unchanged'

    $script:Config_MediaServer = 'plex'
    Assert-Equal 'Fargo - S01E03.mp4' (Get-SeriesEpisodeFileName -Title 'Fargo' -Season 1 -Episode 3 -Extension '.mp4') 'Plex episode: Fargo - S01E03.mp4'
    Assert-Equal 'Fargo - S01E03.mp4' (Get-SeriesEpisodeFileName -Title 'Fargo' -Season 0 -Episode 3 -Extension '.mp4') 'Plex, no -Season: still S01'
    Assert-Equal 'Fargo - S00E02.mkv' (Get-SeriesSpecialFileName -Title 'Fargo' -Special 2 -Extension '.mkv') 'Plex special: Fargo - S00E02.mkv'
    Assert-Equal 'Fargo-S02-D1-Extra01.mkv' (Get-SeriesExtraFileName -Title 'Fargo' -Season 2 -Extra 1 -Extension '.mkv' -Disc 1) 'extras keep their name (Plex does not parse names in Other)'

    $pat = Get-SeriesNamePattern -Title 'Fargo' -Season 2
    Assert-True (('Fargo - S02E05.mp4' -match $pat) -and [int]$Matches['ep'] -eq 5) 'a Plex episode counts as already renamed (episode 5 reserved)'
    Assert-True (('Fargo-S02-E06.mkv' -match $pat) -and [int]$Matches['ep'] -eq 6) 'so does a Jellyfin episode, whatever the setting'
    Assert-True (('Fargo-S02-D1-Extra03.mp4' -match $pat) -and [int]$Matches['extra'] -eq 3 -and -not $Matches['ep']) 'extras still reserve their numbers'
    Assert-True ('Fargo - S03E01.mp4' -notmatch $pat -and 'Fargo - S02E05 - Pilot.mp4' -notmatch $pat -and 'title_t00.mp4' -notmatch $pat) 'other seasons, other shapes and raw files are not "already renamed"'

    $sp = Join-Path $tempRoot 'Fargo\Specials'
    New-Item -ItemType Directory -Path $sp -Force | Out-Null
    'Fargo - S00E01.mp4', 'Fargo-S00-E02.mkv', 'Other - S00E03.mp4' | ForEach-Object { Set-Content -LiteralPath (Join-Path $sp $_) -Value 'x' }
    Assert-Equal '1,2' ((Get-SeriesTakenSpecials -SpecialsDir $sp -Title 'Fargo' | Sort-Object) -join ',') 'taken specials read both forms, this title only'

    $season = Join-Path $tempRoot 'Fargo\Season 1'
    New-Item -ItemType Directory -Path (Join-Path $season 'Disc1') -Force | Out-Null
    1..4 | ForEach-Object { Set-Content -LiteralPath (Join-Path $season ('Disc1\Fargo - S01E{0:D2}.mp4' -f $_)) -Value 'x' }
    Assert-Equal 5 (Get-SuggestedStartEpisode -SeasonDir $season -Disc 2) 'Disc 2 start suggestion reads Plex names on Disc 1'

    $classification = [pscustomobject]@{ Items = @(
        [pscustomobject]@{ Name = 'a.mp4'; Kind = 'Episode'; EpisodeNumber = 1; ExtraNumber = 0; SpecialNumber = 0; Label = ''; Source = 'Median'; DurationSec = 1800; ExpectedSec = $null; Note = '' },
        [pscustomobject]@{ Name = 'b.mp4'; Kind = 'Special'; EpisodeNumber = 0; ExtraNumber = 0; SpecialNumber = 1; Label = ''; Source = 'Median'; DurationSec = 6000; ExpectedSec = $null; Note = '' },
        [pscustomobject]@{ Name = 'c.mp4'; Kind = 'Extra'; EpisodeNumber = 0; ExtraNumber = 1; SpecialNumber = 0; Label = ''; Source = 'Median'; DurationSec = 300; ExpectedSec = $null; Note = '' }) }
    $planDir = Join-Path $tempRoot 'Plan\Season 1\Disc1'
    $plan = New-SeriesRenamePlan -Classification $classification -Title 'Fargo' -Season 1 -Directory $planDir
    Assert-Equal 'Fargo - S01E01.mp4|..\..\Specials\Fargo - S00E01.mp4|..\..\Other\Fargo-S01-D1-Extra01.mp4' (($plan | ForEach-Object { $_.NewName }) -join '|') 'Plex plan: episode in the disc folder, special in Specials, extra in Other'
    $script:Config_MediaServer = 'jellyfin'
    $plan = New-SeriesRenamePlan -Classification $classification -Title 'Fargo' -Season 1 -Directory $planDir
    Assert-Equal 'Fargo-S01-E01.mp4|..\..\Season 0\Fargo-S00-E01.mp4|..\..\Season 0\Fargo-S01-D1-Extra01.mp4' (($plan | ForEach-Object { $_.NewName }) -join '|') 'Jellyfin plan unchanged: special and extra in Season 0'

    # ---------------------------------------------------------------------------
    Write-Host "`nSeries: Plex rename end to end, re-run under Jellyfin, undo" -ForegroundColor Cyan

    $script:Config_TmdbApiKey = $null
    $env:TMDB_API_KEY = $null
    function Invoke-TMDbRequest { param([string]$Url) throw 'must not be called without a key' }

    $script:Config_MediaServer = 'plex'
    $seriesDir = Join-Path $tempRoot 'Series\Silicon Valley'
    $discDir = Join-Path $seriesDir 'Season 2\Disc1'
    New-Item -ItemType Directory -Path $discDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mkv', 'title_t03.mp4' | ForEach-Object { Set-Content -Path (Join-Path $discDir $_) -Value $_ }
    $durations = @{ 'title_t00.mp4' = 1770; 'title_t01.mp4' = 1800; 'title_t02.mkv' = 1760; 'title_t03.mp4' = 300 }
    $getDuration = { param($path) $durations[(Split-Path -Leaf $path)] }.GetNewClosure()

    $result = Invoke-SeriesEpisodeRename -Directory $discDir -Title 'Silicon Valley' -Season 2 -StartEpisode 1 `
        -UndoScriptSource $undoPath -ReadInput (New-InputQueue @('')) -GetDuration $getDuration 6>$null
    Assert-Equal 4 $result.Renamed 'four files renamed'
    Assert-Equal 'Other\Silicon Valley-S02-D1-Extra01.mp4,Season 2\Disc1\Silicon Valley - S02E01.mp4,Season 2\Disc1\Silicon Valley - S02E02.mp4,Season 2\Disc1\Silicon Valley - S02E03.mkv' (Get-RelNames $seriesDir) 'Plex: episodes as "Title - S02E0N", the extra in the series-level Other folder'
    $manifest = Join-Path $discDir 'rename-manifest.csv'
    $extraRow = @(Import-Csv -LiteralPath $manifest | Where-Object { $_.Kind -eq 'Extra' })
    Assert-Equal '..\..\Other\Silicon Valley-S02-D1-Extra01.mp4' $extraRow[0].NewName 'manifest records the extra relative to the disc folder'

    $script:Config_MediaServer = 'jellyfin'
    $again = Invoke-SeriesEpisodeRename -Directory $discDir -Title 'Silicon Valley' -Season 2 -StartEpisode 1 `
        -UndoScriptSource $undoPath -ReadInput (New-InputQueue @()) -GetDuration $getDuration 6>$null
    Assert-Equal 0 $again.Renamed 'switching to Jellyfin later never renames the Plex-named files again'

    $undo = & $undoPath -ManifestPath $manifest 6>$null 3>$null
    Assert-Equal 4 $undo.Restored 'undo-rename.ps1 accepts ..\..\Other\ rows and restores all four'
    Assert-Equal 'Season 2\Disc1\title_t00.mp4,Season 2\Disc1\title_t01.mp4,Season 2\Disc1\title_t02.mkv,Season 2\Disc1\title_t03.mp4' (Get-RelNames $seriesDir) 'original names back'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $seriesDir 'Other'))) 'the emptied Other folder is removed'

    $undoText = Get-Content -LiteralPath $undoPath -Raw
    Assert-True ($undoText.Contains('(?:Season 0|Specials|Other)')) 'undo whitelist names exactly Season 0, Specials and Other'
    $retroText = Get-Content -LiteralPath $retroPath -Raw
    Assert-True ($retroText.Contains("'(?:-|S\d+)E(\d+)\.[^.\\]+$'") -and $retroText.Contains("'(?:-S00-|S00)E(\d+)\.[^.\\]+$'")) 'rename-series.ps1 carries numbering forward from either form'
    Assert-Equal '07' ([regex]::Match('..\x - S02E07.mp4', '(?:-|S\d+)E(\d+)\.[^.\\]+$').Groups[1].Value) '... Plex episode number read'
    Assert-Equal '07' ([regex]::Match('x-S02-E07.mp4', '(?:-|S\d+)E(\d+)\.[^.\\]+$').Groups[1].Value) '... Jellyfin episode number read'

    # ---------------------------------------------------------------------------
    Write-Host "`nGenre series episode names" -ForegroundColor Cyan

    $script:Config_MediaServer = 'jellyfin'
    Assert-Equal 'The Blues-S01E04.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'S01E04' -EpisodeName '' -Extension '.mp4') 'Jellyfin unnamed episode unchanged'
    $script:Config_MediaServer = 'plex'
    Assert-Equal 'The Blues - S01E04.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'S01E04' -EpisodeName '' -Extension '.mp4') 'Plex unnamed episode: " - " separator'
    Assert-Equal 'The Blues - S01E04 - Warming.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'S01E04' -EpisodeName 'Warming' -Extension '.mp4') 'named episode is the same for both'

    # ---------------------------------------------------------------------------
    Write-Host "`nsetup.ps1 and the sample config" -ForegroundColor Cyan

    $setupText = Get-Content -LiteralPath $setupPath -Raw
    Assert-True ($setupText.Contains('[8/8] Media server')) 'setup.ps1 asks for the media server as step 8'
    Assert-True ($setupText -match 'mediaServer = \$mediaServer') 'setup.ps1 saves mediaServer'
    Assert-True ($setupText -match "\`$mediaServer = 'plex'" -and $setupText -match "\`$mediaServer = 'jellyfin'") 'setup.ps1 offers both servers'
    $sample = Get-Content -LiteralPath $samplePath -Raw | ConvertFrom-Json
    Assert-Equal 'jellyfin' $sample.mediaServer 'sample config defaults to Jellyfin'
    $nonAscii = Select-String -LiteralPath $configPath -Pattern '[^\x00-\x7F]'
    Assert-True (-not $nonAscii) 'Load-Config.ps1 stays ASCII (Windows PowerShell reads BOM-less files as ANSI)'
} finally {
    $script:Config_MediaServer = 'jellyfin'
    $env:TMDB_API_KEY = $savedEnvKey
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
