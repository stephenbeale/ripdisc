#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for plain -Series episode naming: <Title>-S##-E## names, extras classification
    (TMDb runtimes and the median fallback), the Disc 2+ start-episode prompt, the
    rename manifest and undo-rename.ps1.

.DESCRIPTION
    Dot-sources the real SeriesEpisodes.ps1 (the file rip-disc.ps1 and continue-rip.ps1
    both load) and lifts Get-ContinueRipCommand out of rip-disc.ps1 with the AST parser,
    so nothing here is a reimplementation. TMDb is mocked by replacing the single
    Invoke-TMDbRequest seam - no network. Filesystem cases use a temp directory with
    empty placeholder video files, removed afterwards. No disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-SeriesEpisodeRename.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$sharedPath = Join-Path $repoRoot 'SeriesEpisodes.ps1'
$undoPath = Join-Path $repoRoot 'undo-rename.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($sharedPath, $undoPath, $ripDiscPath, $continuePath)) {
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

. $sharedPath
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-ContinueRipCommand')))

# The scripts' own Write-Log appends to the session log; collect it here instead.
$script:LogLines = New-Object System.Collections.Generic.List[string]
function Write-Log { param([string]$Message) $script:LogLines.Add($Message) }

# Scripted console input: each call returns the next queued answer ($null when empty).
function New-InputQueue {
    param([object[]]$Answers)
    $queue = New-Object System.Collections.Queue
    foreach ($a in $Answers) { $queue.Enqueue($a) }
    return { param($prompt) if ($queue.Count -gt 0) { $queue.Dequeue() } else { $null } }.GetNewClosure()
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

function Get-Kinds { param($Classification) ($Classification.Items | ForEach-Object { if ($_.Kind -eq 'Episode') { "E$($_.EpisodeNumber)" } else { "X$($_.ExtraNumber)" } }) -join ',' }
function New-Titles { param([double[]]$Minutes) $i = 0; foreach ($m in $Minutes) { [pscustomobject]@{ Name = ('title_t{0:D2}.mkv' -f $i); DurationSec = if ($m -gt 0) { $m * 60 } else { $null } }; $i++ } }
# Video files under a Disc folder, as paths relative to it (extras live in extras\).
function Get-RelNames {
    param([string]$Dir)
    $root = (Resolve-Path -LiteralPath $Dir).ProviderPath.TrimEnd('\') + '\'
    (Get-ChildItem -LiteralPath $Dir -File -Recurse | Where-Object { $_.Extension -in '.mp4', '.mkv' } |
        ForEach-Object { $_.FullName.Substring($root.Length) } | Sort-Object) -join ','
}
function New-TmdbEpisodes { param([int[]]$Runtimes) $n = 1; foreach ($r in $Runtimes) { [pscustomobject]@{ Number = $n; Name = "Ep $n"; Runtime = $r }; $n++ } }

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-series-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nFile name format" -ForegroundColor Cyan

    Assert-Equal 'Silicon Valley-S02-E01.mp4' (Get-SeriesEpisodeFileName -Title 'Silicon Valley' -Season 2 -Episode 1 -Extension '.mp4') 'season 2 episode 1 -> Silicon Valley-S02-E01.mp4'
    Assert-Equal 'Silicon Valley-S02-E10.mkv' (Get-SeriesEpisodeFileName -Title 'Silicon Valley' -Season 2 -Episode 10 -Extension '.mkv') 'the real extension is kept (.mkv is not changed to .mp4)'
    Assert-Equal 'Fargo-S01-E03.mp4' (Get-SeriesEpisodeFileName -Title 'Fargo' -Season 0 -Episode 3 -Extension '.mp4') 'no -Season falls back to S01'
    Assert-Equal 'Doctor Who-S12-E104.mp4' (Get-SeriesEpisodeFileName -Title 'Doctor Who' -Season 12 -Episode 104 -Extension '.mp4') 'episode numbers past 99 widen rather than truncate'
    Assert-Equal 'Silicon Valley-S02-Extra01.mkv' (Get-SeriesExtraFileName -Title 'Silicon Valley' -Season 2 -Extra 1 -Extension '.mkv') 'extras are named <Title>-S##-Extra##'

    $pattern = Get-SeriesNamePattern -Title 'W.' -Season 2
    Assert-True ('W.-S02-E05.mp4' -match $pattern -and $Matches['ep'] -eq '05') 'already-renamed pattern recognises an episode and captures its number'
    Assert-True ('W.-S02-Extra02.mkv' -match $pattern -and $Matches['extra'] -eq '02') 'already-renamed pattern recognises an extra'
    Assert-True (-not ('WX-S02-E05.mp4' -match $pattern)) 'title is regex-escaped (the "." in "W." is literal)'
    Assert-True (-not ('W.-S03-E05.mp4' -match $pattern)) 'a different season is not treated as already renamed'
    Assert-True (-not ('title_t00.mp4' -match $pattern)) 'MakeMKV output names are not treated as already renamed'

    # ---------------------------------------------------------------------------
    Write-Host "`nStart-episode prompt: when it is (and is not) asked" -ForegroundColor Cyan

    Assert-True (Test-ShouldPromptStartEpisode -Series -Disc 2) 'plain -Series, Disc 2, no -StartEpisode -> prompt'
    Assert-True (-not (Test-ShouldPromptStartEpisode -Series -Disc 1)) 'Disc 1 never prompts'
    Assert-True (-not (Test-ShouldPromptStartEpisode -Series -Disc 3 -StartEpisodeExplicit)) 'explicit -StartEpisode (or an earlier answer) never prompts'
    Assert-True (-not (Test-ShouldPromptStartEpisode -Series -GenreSeries -Disc 2)) 'genre series keeps its own auto-detection - no prompt'
    Assert-True (-not (Test-ShouldPromptStartEpisode -Series -Extras -Disc 2)) '-Extras disc has no episodes - no prompt'
    Assert-True (-not (Test-ShouldPromptStartEpisode -Disc 2)) 'movie mode never prompts'

    Write-Host "`nStart-episode prompt: answers" -ForegroundColor Cyan

    Assert-Equal 5 (Read-StartEpisode -Disc 2 -Suggested 5 -ReadInput (New-InputQueue @(''))) 'Enter accepts the suggested start'
    Assert-Equal 9 (Read-StartEpisode -Disc 2 -Suggested 5 -ReadInput (New-InputQueue @('9'))) 'a typed number overrides the suggestion'
    Assert-Equal 7 (Read-StartEpisode -Disc 2 -Suggested $null -ReadInput (New-InputQueue @('', 'abc', '0', '7'))) 'blank (no suggestion), text and 0 are rejected until a valid number arrives'
    Assert-Equal 5 (Read-StartEpisode -Disc 2 -Suggested 5 -ReadInput (New-InputQueue @())) 'no console input available -> takes the suggestion'
    $threw = $false
    try { Read-StartEpisode -Disc 2 -Suggested $null -ReadInput (New-InputQueue @()) | Out-Null } catch { $threw = $true }
    Assert-True $threw 'no console input and no suggestion -> stops with a clear error instead of looping'

    Write-Host "`nStart-episode suggestion from earlier discs" -ForegroundColor Cyan

    $seasonDir = Join-Path $tempRoot 'Series\Show\Season 2'
    foreach ($d in 'Disc1', 'Disc2', 'Disc3') { New-Item -ItemType Directory -Path (Join-Path $seasonDir $d) -Force | Out-Null }
    'Show-S02-E01.mp4', 'Show-S02-E02.mp4', 'Show-S02-E03.mp4', 'Show-S02-Extra01.mp4' | ForEach-Object { New-Item -ItemType File -Path (Join-Path $seasonDir "Disc1\$_") | Out-Null }
    'Show-S02-E40.mp4' | ForEach-Object { New-Item -ItemType File -Path (Join-Path $seasonDir "Disc3\$_") | Out-Null }
    Assert-Equal 4 (Get-SuggestedStartEpisode -SeasonDir $seasonDir -Disc 2) 'Disc 2 suggestion is Disc 1''s highest episode + 1 (extras and later discs ignored)'
    [pscustomobject]@{ OriginalName = 'title_t03.mp4'; NewName = 'Show-S02-E04.mp4'; Kind = 'Episode'; OriginalPath = ''; NewPath = ''; Timestamp = '' } |
        Export-Csv -LiteralPath (Join-Path $seasonDir 'Disc1\rename-manifest.csv') -NoTypeInformation
    Assert-Equal 5 (Get-SuggestedStartEpisode -SeasonDir $seasonDir -Disc 2) 'manifest rows on an earlier disc count too'
    Assert-Equal $null (Get-SuggestedStartEpisode -SeasonDir $seasonDir -Disc 1) 'Disc 1 has no earlier discs -> no suggestion'
    Assert-Equal $null (Get-SuggestedStartEpisode -SeasonDir (Join-Path $tempRoot 'nope') -Disc 2) 'missing season folder -> no suggestion'

    Write-Host "`nResumed rips carry the chosen start episode" -ForegroundColor Cyan

    $cmd = Get-ContinueRipCommand -Title 'Show' -RemainingStepNumbers @(3, 4) -Series -Season 2 -Disc 2 -StartEpisode 1 -EpisodeNames @() -StartEpisodeExplicit
    Assert-Equal '.\continue-rip.ps1 -title "Show" -FromStep organize -Series -Season 2 -Disc 2 -StartEpisode 1' $cmd '-StartEpisode 1 is still printed when it was answered at the prompt, so continue-rip.ps1 does not ask again'
    $cmd = Get-ContinueRipCommand -Title 'Show' -RemainingStepNumbers @(3, 4) -Series -Season 2 -Disc 1 -StartEpisode 1 -EpisodeNames @()
    Assert-Equal '.\continue-rip.ps1 -title "Show" -FromStep organize -Series -Season 2' $cmd 'a defaulted -StartEpisode 1 is still omitted (unchanged behaviour)'
    Assert-True ((Get-Content $continuePath -Raw) -match 'Test-ShouldPromptStartEpisode[^\r\n]*-StartEpisodeExplicit:\$script:StartEpisodeExplicit') 'continue-rip.ps1 skips the prompt when -StartEpisode was passed'

    # ---------------------------------------------------------------------------
    Write-Host "`nClassification without TMDb (median heuristic)" -ForegroundColor Cyan

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 45, 43, 6, 44)
    Assert-Equal 'E1,E2,E3,X1,E4' (Get-Kinds $c) 'a 6-minute title among ~44-minute ones is an extra and does not use up an episode number'
    Assert-Equal 'Median' $c.Method 'method reported as Median when no TMDb data'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 132, 44, 44) -StartEpisode 5
    Assert-Equal 'E5,X1,E6,E7' (Get-Kinds $c) 'play-all title (about the sum of the others) is an extra; numbering starts at -StartEpisode'
    Assert-True ($c.Items[1].Note -match 'play-all') 'play-all title is labelled as such'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 0, 44)
    Assert-Equal 'E1,E2,E3' (Get-Kinds $c) 'a title with no readable duration stays an episode...'
    Assert-True ($c.Items[1].Mismatch) '...but is flagged for checking'

    # 45 is not 70-130% of 22+23+22=67, so it is not a play-all. (With only 22, 23, 45
    # it WOULD be - a double episode equal to the sum of the others is indistinguishable
    # from a play-all by length alone; that is what the confirmation prompt is for.)
    $c = Get-SeriesTitleClassification -Titles (New-Titles 22, 23, 22, 45)
    Assert-Equal 'E1,E2,E3,E4' (Get-Kinds $c) 'long titles (double episodes) are never demoted to extras'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 44, 44, 5) -TakenEpisodes @(1, 2) -TakenExtras @(1)
    Assert-Equal 'E3,E4,E5,X2' (Get-Kinds $c) 'numbers already used in the folder are skipped (re-run after a part-way failure)'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 44, 5) -Overrides @{ 'title_t01.mkv' = 'Extra'; 'title_t02.mkv' = 'Episode' }
    Assert-Equal 'E1,X1,E2' (Get-Kinds $c) 'overrides from the edit prompt always win'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 44, 44) -AllExtras
    Assert-Equal 'X1,X2' (Get-Kinds $c) '-Extras disc: everything is an extra'

    Write-Host "`nClassification with TMDb per-episode runtimes" -ForegroundColor Cyan

    $tmdb = @(New-TmdbEpisodes 30, 28, 30, 30, 30, 29, 30, 30)
    $c = Get-SeriesTitleClassification -Titles (New-Titles 29.5, 27, 30.2, 29) -StartEpisode 1 -TmdbEpisodes $tmdb -TmdbEpisodeCount 8
    Assert-Equal 'E1,E2,E3,E4' (Get-Kinds $c) 'titles within tolerance of E1-E4 runtimes are episodes'
    Assert-Equal 'TMDb' $c.Method 'method reported as TMDb'
    Assert-Equal (28 * 60) $c.Items[1].ExpectedSec 'expected runtime is the one for the episode the title maps to (E02 = 28 min)'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 30, 12, 29, 30) -StartEpisode 5 -TmdbEpisodes $tmdb -TmdbEpisodeCount 8
    Assert-Equal 'E5,X1,E6,E7' (Get-Kinds $c) 'a title far shorter than the next episode''s runtime is an extra and E06 goes to the next title'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 30, 22) -StartEpisode 1 -TmdbEpisodes $tmdb -TmdbEpisodeCount 8
    Assert-Equal 'E1,E2' (Get-Kinds $c) 'a title outside tolerance but not clearly short stays an episode...'
    Assert-True ($c.Items[1].Mismatch -and $c.Items[1].Note -match 'differs from TMDb E02') '...flagged with which TMDb episode it differs from'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 30, 30, 30) -StartEpisode 7 -TmdbEpisodes $tmdb -TmdbEpisodeCount 8
    Assert-Equal 'E7,E8,E9' (Get-Kinds $c) 'numbering past the season is not silently dropped...'
    Assert-True (@($c.Warnings | Where-Object { $_ -match 'only 2 episode\(s\) left' }).Count -eq 1) '...and warns that the disc has more episode-length titles than episodes left'
    Assert-True ($c.Items[2].Note -match "past TMDb's 8 episodes") 'the title beyond the last TMDb episode is flagged'

    $c = Get-SeriesTitleClassification -Titles (New-Titles 30, 30) -TmdbEpisodes @(New-TmdbEpisodes 0, 0) -TmdbEpisodeCount 2
    Assert-Equal 'Median' $c.Method 'TMDb episodes without runtimes fall back to the median heuristic'

    Write-Host "`nHandBrake scan duration parsing" -ForegroundColor Cyan

    Assert-Equal 2652 (ConvertFrom-HandBrakeScanDuration -Lines @('+ title 1:', '  + duration: 00:44:12', '  + size: 720x576')) 'parses "+ duration: 00:44:12" to seconds'
    Assert-Equal $null (ConvertFrom-HandBrakeScanDuration -Lines @('no title found')) 'returns $null when there is no duration line'

    # ---------------------------------------------------------------------------
    Write-Host "`nTMDb season lookup (mocked, fail-soft)" -ForegroundColor Cyan

    $script:Config_TmdbApiKey = 'test-key'
    $script:TmdbCalls = New-Object System.Collections.Generic.List[string]
    function Invoke-TMDbRequest {
        param([string]$Url)
        $script:TmdbCalls.Add($Url)
        if ($Url -match '/search/tv') {
            return [pscustomobject]@{ results = @(
                [pscustomobject]@{ id = 111; name = 'Silicon Valley Rebels'; first_air_date = '2001-01-01' },
                [pscustomobject]@{ id = 60573; name = 'Silicon Valley'; first_air_date = '2014-04-06' }
            ) }
        }
        if ($Url -match '/tv/60573/season/2\?') {
            return [pscustomobject]@{ episodes = @(
                [pscustomobject]@{ episode_number = 1; name = 'Sand Hill Shuffle'; runtime = 29 },
                [pscustomobject]@{ episode_number = 2; name = 'Runaway Devaluation'; runtime = 30 }
            ) }
        }
        throw "unexpected URL $Url"
    }

    $season = Get-SeriesTmdbSeason -Title 'Silicon Valley' -Season 2
    Assert-True ($null -ne $season) 'season details returned'
    Assert-Equal 60573 $season.TvId 'the exact name match is preferred over the first search result'
    Assert-Equal 2 $season.EpisodeCount 'episode count comes from the season endpoint'
    Assert-Equal 29 $season.Episodes[0].Runtime 'per-episode runtime is kept'

    $script:TmdbCalls.Clear()
    $season = Get-SeriesTmdbSeason -Title 'Anything' -Season 2 -KnownTvId 60573
    Assert-True ($null -ne $season -and $script:TmdbCalls.Count -eq 1 -and $script:TmdbCalls[0] -notmatch '/search/') 'a TMDb id already found by auto-discovery skips the search call'

    function Invoke-TMDbRequest { param([string]$Url) throw 'The remote name could not be resolved' }
    Assert-Equal $null (Get-SeriesTmdbSeason -Title 'Silicon Valley' -Season 2) 'offline / network error -> $null, no exception'

    function Invoke-TMDbRequest { param([string]$Url) return [pscustomobject]@{ results = @() } }
    Assert-Equal $null (Get-SeriesTmdbSeason -Title 'Nothing Matches' -Season 2) 'no TV match -> $null'

    $script:Config_TmdbApiKey = $null
    $savedEnvKey = $env:TMDB_API_KEY
    $env:TMDB_API_KEY = $null
    function Invoke-TMDbRequest { param([string]$Url) throw 'must not be called without a key' }
    Assert-Equal $null (Get-SeriesTmdbSeason -Title 'Silicon Valley' -Season 2) 'no API key -> $null without any network call'
    $env:TMDB_API_KEY = $savedEnvKey

    # ---------------------------------------------------------------------------
    Write-Host "`nRename end to end: manifest, names, undo script" -ForegroundColor Cyan

    $discDir = Join-Path $tempRoot 'Series\Silicon Valley\Season 2\Disc1'
    New-Item -ItemType Directory -Path $discDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mkv', 'title_t03.mp4' | ForEach-Object { Set-Content -Path (Join-Path $discDir $_) -Value $_ }
    $durations = @{ 'title_t00.mp4' = 1770; 'title_t01.mp4' = 1800; 'title_t02.mkv' = 1760; 'title_t03.mp4' = 300 }
    $getDuration = { param($path) $durations[(Split-Path -Leaf $path)] }.GetNewClosure()

    $result = Invoke-SeriesEpisodeRename -Directory $discDir -Title 'Silicon Valley' -Season 2 -StartEpisode 1 `
        -UndoScriptSource $undoPath -ReadInput (New-InputQueue @('')) -GetDuration $getDuration
    $names = Get-RelNames $discDir
    Assert-Equal 'extras\Silicon Valley-S02-Extra01.mp4,Silicon Valley-S02-E01.mp4,Silicon Valley-S02-E02.mp4,Silicon Valley-S02-E03.mkv' $names 'episodes renamed in MakeMKV order (extension kept); the short extra moved into the Disc folder''s extras subfolder'
    Assert-Equal 4 $result.Renamed 'four files renamed'

    $manifestPath = Join-Path $discDir 'rename-manifest.csv'
    Assert-True (Test-Path $manifestPath) 'rename-manifest.csv written in the Disc folder'
    $rows = @(Import-Csv $manifestPath)
    Assert-Equal 4 $rows.Count 'one manifest row per renamed file'
    Assert-Equal 'OriginalName,NewName,Kind,OriginalPath,NewPath,Timestamp' (($rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'manifest columns'
    Assert-Equal 'title_t03.mp4|extras\Silicon Valley-S02-Extra01.mp4|Extra' ("{0}|{1}|{2}" -f $rows[3].OriginalName, $rows[3].NewName, $rows[3].Kind) 'extras are in the manifest with their path relative to the Disc folder'
    Assert-Equal (Join-Path $discDir 'extras\Silicon Valley-S02-Extra01.mp4') $rows[3].NewPath 'manifest NewPath is the full path inside extras'
    Assert-True ($rows[0].Timestamp -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') 'manifest rows carry a timestamp'
    Assert-True (Test-Path (Join-Path $discDir 'undo-rename.ps1')) 'undo-rename.ps1 copied next to the manifest'
    Assert-True (@($script:LogLines | Where-Object { $_ -like 'Renamed (Episode): title_t00.mp4 -> Silicon Valley-S02-E01.mp4' }).Count -eq 1) 'each rename is written to the session log'

    Write-Host "`nUndo" -ForegroundColor Cyan

    $copiedUndo = Join-Path $discDir 'undo-rename.ps1'
    $whatIf = & $copiedUndo -WhatIf 3>$null 6>$null
    Assert-True (Test-Path (Join-Path $discDir 'Silicon Valley-S02-E01.mp4')) '-WhatIf changes nothing'

    # A file now sitting at an original name must not be overwritten by the undo.
    Set-Content -Path (Join-Path $discDir 'title_t01.mp4') -Value 'squatter'
    $undo = & $copiedUndo 3>$null 6>$null
    Assert-Equal 3 $undo.Restored 'undo (copied script, default manifest path) restores the other three'
    Assert-Equal 1 $undo.Skipped 'name collision is skipped, not overwritten'
    Assert-Equal 'squatter' ((Get-Content (Join-Path $discDir 'title_t01.mp4')) -join '') 'the colliding file is untouched'
    Assert-True (Test-Path (Join-Path $discDir 'Silicon Valley-S02-E02.mp4')) 'the renamed file that collided keeps its new name'
    Assert-True ((Test-Path (Join-Path $discDir 'title_t00.mp4')) -and (Test-Path (Join-Path $discDir 'title_t02.mkv')) -and (Test-Path (Join-Path $discDir 'title_t03.mp4'))) 'restored files are back under their original names'

    $undo = & $undoPath -ManifestPath $manifestPath 3>$null 6>$null
    Assert-Equal 0 $undo.Restored 'running undo again (repo script, -ManifestPath) restores nothing new...'
    Assert-Equal 4 $undo.Skipped '...and skips missing/colliding files with warnings instead of failing'

    Write-Host "`nManifest survives a failure part-way through" -ForegroundColor Cyan

    $failDir = Join-Path $tempRoot 'Series\Fargo\Disc2'
    New-Item -ItemType Directory -Path $failDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mp4' | ForEach-Object { Set-Content -Path (Join-Path $failDir $_) -Value $_ }
    $titles = @('title_t00.mp4', 'title_t01.mp4', 'title_t02.mp4' | ForEach-Object { [pscustomobject]@{ Name = $_; DurationSec = 2640 } })
    $plan = New-SeriesRenamePlan -Classification (Get-SeriesTitleClassification -Titles $titles -StartEpisode 5) -Title 'Fargo' -Season 0 -Directory $failDir
    $failManifest = Join-Path $failDir 'rename-manifest.csv'
    $null = Write-RenameManifest -Plan $plan -ManifestPath $failManifest
    $lock = [System.IO.File]::Open((Join-Path $failDir 'title_t01.mp4'), 'Open', 'Read', 'None')
    $threw = $false
    try { Invoke-SeriesRenamePlan -Plan $plan -MaxRetries 1 -RetryDelaySec 0 6>$null | Out-Null } catch { $threw = $true } finally { $lock.Close() }
    Assert-True $threw 'a locked file stops the rename part-way through'
    Assert-Equal 3 @(Import-Csv $failManifest).Count 'the manifest already lists every planned rename'
    Assert-True ((Test-Path (Join-Path $failDir 'Fargo-S01-E05.mp4')) -and (Test-Path (Join-Path $failDir 'title_t01.mp4'))) 'first file renamed, the locked one untouched (no-season -> S01)'

    $result = Invoke-SeriesEpisodeRename -Directory $failDir -Title 'Fargo' -Season 0 -StartEpisode 5 `
        -ReadInput (New-InputQueue @('')) -GetDuration { param($p) 2640 } 6>$null
    $names = (Get-ChildItem -LiteralPath $failDir -Filter '*.mp4' | Sort-Object Name | ForEach-Object { $_.Name }) -join ','
    Assert-Equal 'Fargo-S01-E05.mp4,Fargo-S01-E06.mp4,Fargo-S01-E07.mp4' $names 're-running picks up where it stopped without renumbering the finished file'
    Assert-Equal 5 @(Import-Csv $failManifest).Count 'the re-run appends to the existing manifest'
    $undo = & $undoPath -ManifestPath $failManifest 3>$null 6>$null
    $names = (Get-ChildItem -LiteralPath $failDir -Filter '*.mp4' | Sort-Object Name | ForEach-Object { $_.Name }) -join ','
    Assert-Equal 'title_t00.mp4,title_t01.mp4,title_t02.mp4' $names 'undo restores every file, warning on the stale rows from the failed attempt'

    Write-Host "`nNo overwrite, decline and edit at the confirmation prompt" -ForegroundColor Cyan

    $promptDir = Join-Path $tempRoot 'Series\Prompt\Disc1'
    New-Item -ItemType Directory -Path $promptDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mp4' | ForEach-Object { Set-Content -Path (Join-Path $promptDir $_) -Value $_ }

    $result = Invoke-SeriesEpisodeRename -Directory $promptDir -Title 'Prompt' -Season 1 -ReadInput (New-InputQueue @('n')) -GetDuration { param($p) 2640 } 6>$null
    Assert-True ($result.Declined -and -not (Test-Path (Join-Path $promptDir 'rename-manifest.csv'))) '"n" leaves every file alone and writes no manifest'
    Assert-True (Test-Path (Join-Path $promptDir 'title_t00.mp4')) 'files keep their MakeMKV names after declining'

    $result = Invoke-SeriesEpisodeRename -Directory $promptDir -Title 'Prompt' -Season 1 -ReadInput (New-InputQueue @('e', '2', '')) -GetDuration { param($p) 2640 } 6>$null
    $names = Get-RelNames $promptDir
    Assert-Equal 'extras\Prompt-S01-Extra01.mp4,Prompt-S01-E01.mp4,Prompt-S01-E02.mp4' $names '"e" then row 2 switches that title to an extra and renumbers the rest'

    $collideDir = Join-Path $tempRoot 'Series\Collide\Disc1'
    New-Item -ItemType Directory -Path $collideDir -Force | Out-Null
    Set-Content -Path (Join-Path $collideDir 'title_t00.mp4') -Value 'new'
    $titles = @([pscustomobject]@{ Name = 'title_t00.mp4'; DurationSec = 2640 })
    $classification = Get-SeriesTitleClassification -Titles $titles
    Set-Content -Path (Join-Path $collideDir 'Collide-S01-E01.mp4') -Value 'existing'
    $plan = New-SeriesRenamePlan -Classification $classification -Title 'Collide' -Season 1 -Directory $collideDir
    Assert-True ($plan[0].Skip) 'an existing target is marked Skip in the plan'
    Assert-Equal 0 (Write-RenameManifest -Plan $plan -ManifestPath (Join-Path $collideDir 'rename-manifest.csv')) 'skipped renames are not written to the manifest'
    $outcome = Invoke-SeriesRenamePlan -Plan $plan 6>$null
    Assert-Equal 1 $outcome.Skipped 'the rename is skipped...'
    Assert-Equal 'existing' ((Get-Content (Join-Path $collideDir 'Collide-S01-E01.mp4')) -join '') '...and the existing file is never overwritten'

    # ---------------------------------------------------------------------------
    Write-Host "`nExtras subfolder (DiscN\extras, same folder name as movie rips)" -ForegroundColor Cyan

    $noExtrasDir = Join-Path $tempRoot 'Series\NoExtras\Season 1\Disc1'
    New-Item -ItemType Directory -Path $noExtrasDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4' | ForEach-Object { Set-Content -Path (Join-Path $noExtrasDir $_) -Value $_ }
    $null = Invoke-SeriesEpisodeRename -Directory $noExtrasDir -Title 'NoExtras' -Season 1 -ReadInput (New-InputQueue @('')) -GetDuration { param($p) 2640 } 6>$null
    Assert-True (-not (Test-Path (Join-Path $noExtrasDir 'extras'))) 'a disc with no extras gets no empty extras folder'

    $exDir = Join-Path $tempRoot 'Series\Ex\Season 1\Disc2'
    New-Item -ItemType Directory -Path $exDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mp4', 'title_t03.mp4', 'title_t04.mkv' | ForEach-Object { Set-Content -Path (Join-Path $exDir $_) -Value $_ }
    $exDur = @{ 'title_t00.mp4' = 2640; 'title_t01.mp4' = 2600; 'title_t02.mp4' = 2620; 'title_t03.mp4' = 300; 'title_t04.mkv' = 240 }
    $exGet = { param($path) $exDur[(Split-Path -Leaf $path)] }.GetNewClosure()
    $null = Invoke-SeriesEpisodeRename -Directory $exDir -Title 'Ex' -Season 1 -StartEpisode 5 -UndoScriptSource $undoPath `
        -ReadInput (New-InputQueue @('')) -GetDuration $exGet 6>$null
    Assert-Equal 'Ex-S01-E05.mp4,Ex-S01-E06.mp4,Ex-S01-E07.mp4,extras\Ex-S01-Extra01.mp4,extras\Ex-S01-Extra02.mkv' (Get-RelNames $exDir) 'episodes stay in DiscN; both extras moved to DiscN\extras (lowercase, like movie rips)'
    Assert-True ((Test-Path (Join-Path $exDir 'rename-manifest.csv')) -and -not (Test-Path (Join-Path $exDir 'extras\rename-manifest.csv'))) 'one manifest, in the Disc folder (not inside extras)'
    $exRows = @(Import-Csv (Join-Path $exDir 'rename-manifest.csv'))
    Assert-Equal 'extras\Ex-S01-Extra02.mkv' $exRows[4].NewName 'manifest NewName is the path relative to the Disc folder'
    Assert-Equal 8 (Get-SuggestedStartEpisode -SeasonDir (Split-Path $exDir -Parent) -Disc 3) 'next-disc suggestion still counts episodes only (extras in the subfolder are ignored)'

    # A later re-run (e.g. a forgotten file) must not reuse Extra01/02 from the subfolder.
    Set-Content -Path (Join-Path $exDir 'title_t09.mp4') -Value 'late'
    $exDur['title_t09.mp4'] = 200
    $null = Invoke-SeriesEpisodeRename -Directory $exDir -Title 'Ex' -Season 1 -StartEpisode 5 -ReadInput (New-InputQueue @('e', '1', '')) -GetDuration $exGet 6>$null
    Assert-True (Test-Path (Join-Path $exDir 'extras\Ex-S01-Extra03.mp4')) 're-run: extras already in the subfolder keep their numbers; the new one becomes Extra03'
    Assert-Equal 6 @(Import-Csv (Join-Path $exDir 'rename-manifest.csv')).Count 're-run appended its row to the same manifest'

    Write-Host "`nUndo moves extras back out of the subfolder" -ForegroundColor Cyan

    $undoResult = & (Join-Path $exDir 'undo-rename.ps1') -WhatIf 3>$null 6>$null
    Assert-True (Test-Path (Join-Path $exDir 'extras\Ex-S01-Extra01.mp4')) '-WhatIf leaves extras where they are'
    $undoResult = & (Join-Path $exDir 'undo-rename.ps1') 3>$null 6>$null
    Assert-Equal 6 $undoResult.Restored 'undo restores all six rows (episodes and extras)'
    Assert-Equal 'title_t00.mp4,title_t01.mp4,title_t02.mp4,title_t03.mp4,title_t04.mkv,title_t09.mp4' (Get-RelNames $exDir) 'every file is back in the Disc folder under its original name'
    Assert-True (-not (Test-Path (Join-Path $exDir 'extras'))) 'the emptied extras folder is removed'

    $keepDir = Join-Path $tempRoot 'Series\Keep\Disc1'
    New-Item -ItemType Directory -Path $keepDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4' | ForEach-Object { Set-Content -Path (Join-Path $keepDir $_) -Value $_ }
    $null = Invoke-SeriesEpisodeRename -Directory $keepDir -Title 'Keep' -Season 1 -ReadInput (New-InputQueue @('e', '2', '')) -GetDuration { param($p) 2640 } 6>$null
    Set-Content -Path (Join-Path $keepDir 'extras\my-notes.txt') -Value 'mine'
    $undoResult = & $undoPath -ManifestPath (Join-Path $keepDir 'rename-manifest.csv') 3>$null 6>$null
    Assert-True ((Test-Path (Join-Path $keepDir 'title_t01.mp4')) -and (Test-Path (Join-Path $keepDir 'extras\my-notes.txt'))) 'an extras folder holding other files is kept after undo'

    $evilDir = Join-Path $tempRoot 'Series\Evil\Disc1'
    New-Item -ItemType Directory -Path (Join-Path $evilDir 'other') -Force | Out-Null
    Set-Content -Path (Join-Path $evilDir 'other\x.mp4') -Value 'x'
    Set-Content -Path (Join-Path $tempRoot 'Series\Evil\up.mp4') -Value 'up'
    @(
        [pscustomobject]@{ OriginalName = 'a.mp4'; NewName = '..\up.mp4'; Kind = 'Extra'; OriginalPath = ''; NewPath = ''; Timestamp = '' }
        [pscustomobject]@{ OriginalName = 'b.mp4'; NewName = 'other\x.mp4'; Kind = 'Extra'; OriginalPath = ''; NewPath = ''; Timestamp = '' }
        [pscustomobject]@{ OriginalName = 'c.mp4'; NewName = 'C:\Windows\x.mp4'; Kind = 'Extra'; OriginalPath = ''; NewPath = ''; Timestamp = '' }
        [pscustomobject]@{ OriginalName = 'd.mp4'; NewName = 'extras\..'; Kind = 'Extra'; OriginalPath = ''; NewPath = ''; Timestamp = '' }
        [pscustomobject]@{ OriginalName = 'extras\e.mp4'; NewName = 'e2.mp4'; Kind = 'Extra'; OriginalPath = ''; NewPath = ''; Timestamp = '' }
    ) | Export-Csv -LiteralPath (Join-Path $evilDir 'rename-manifest.csv') -NoTypeInformation -Encoding UTF8
    $undoResult = & $undoPath -ManifestPath (Join-Path $evilDir 'rename-manifest.csv') 3>$null 6>$null
    Assert-Equal 5 $undoResult.Skipped 'undo refuses rows pointing outside the Disc folder, into other subfolders, rooted, or with a path in OriginalName'
    Assert-True ((Test-Path (Join-Path $evilDir 'other\x.mp4')) -and (Test-Path (Join-Path $tempRoot 'Series\Evil\up.mp4'))) '...and touches none of those files'

    # ---------------------------------------------------------------------------
    Write-Host "`nNon-interactive (-Yes): start episode and confirmation table" -ForegroundColor Cyan

    $auto = Get-AutoStartEpisode -Suggested 9 -Fallback 1 -SuggestedFrom "TheDiscDB's first episode for this disc"
    Assert-True ($auto.Episode -eq 9 -and -not $auto.IsGuess -and $auto.Reason -match 'TheDiscDB' -and $auto.Reason -match '-Yes') 'suggested default is taken, with the reason for the log'
    $auto = Get-AutoStartEpisode -Suggested $null -Fallback 1
    Assert-True ($auto.Episode -eq 1 -and $auto.IsGuess -and $auto.Reason -match 'check') 'no suggestion: falls back to -StartEpisode (1) and flags it as a guess rather than stopping'

    $autoDir = Join-Path $tempRoot 'Series\Auto\Disc1'
    New-Item -ItemType Directory -Path $autoDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4' | ForEach-Object { Set-Content -Path (Join-Path $autoDir $_) -Value $_ }
    $script:LogLines.Clear()
    $neverAsk = { param($p) throw "prompted: $p" }
    $autoResult = Invoke-SeriesEpisodeRename -Directory $autoDir -Title 'Auto' -Season 1 -ReadInput $neverAsk -GetDuration { param($p) 2640 } -AutoAccept 6>$null
    Assert-Equal 2 $autoResult.Renamed '-AutoAccept renames without reading any input'
    Assert-True (@($script:LogLines | Where-Object { $_ -match 'accepted automatically \(-Yes\) - 2 episode\(s\), 0 extra\(s\)' }).Count -eq 1) 'the automatic acceptance is logged with what was chosen'

    $continueText = Get-Content -LiteralPath $continuePath -Raw
    Assert-True ($continueText -match 'Invoke-SeriesEpisodeRename[\s\S]{0,600}-AutoAccept:\$Yes') 'continue-rip.ps1 passes -Yes through as -AutoAccept'
    Assert-True ($continueText -match 'if \(\$Yes\) \{[\s\S]{0,400}Get-AutoStartEpisode[\s\S]{0,600}Write-Log "Start episode chosen automatically') 'continue-rip.ps1 -Yes takes the automatic start episode and logs it instead of prompting'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
