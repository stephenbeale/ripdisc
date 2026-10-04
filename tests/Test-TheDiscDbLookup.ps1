#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the optional TheDiscDB lookup used by plain -Series naming: the disc
    content hash, the GraphQL lookup (match / no match / error / offline / opt-out), lining
    ripped files up with TheDiscDB's titles, classification precedence (TheDiscDB, then
    TMDb, then the median heuristic), named extras, and the -DiscDbHash / -NoDiscDb
    hand-off to continue-rip.ps1.

.DESCRIPTION
    Dot-sources the real SeriesEpisodes.ps1 and lifts Get-ContinueRipCommand out of
    rip-disc.ps1 with the AST parser, so nothing here is a reimplementation. TheDiscDB is
    mocked by replacing the single Invoke-DiscDbRequest seam (the same pattern as
    Invoke-TMDbRequest) - no network. Expected content hashes were computed independently
    (Python: MD5 over little-endian Int64 sizes), not with the code under test. Filesystem
    cases use a temp directory removed afterwards. No disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-TheDiscDbLookup.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$sharedPath = Join-Path $repoRoot 'SeriesEpisodes.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($sharedPath, $ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

function Get-ScriptAst {
    param([string]$ScriptPath)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }
    return $ast
}

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $ast = Get-ScriptAst $ScriptPath
    $fn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

function Get-ScriptParameterNames {
    param([string]$ScriptPath)
    $ast = Get-ScriptAst $ScriptPath
    return @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}

. $sharedPath
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-ContinueRipCommand')))

$script:LogLines = New-Object System.Collections.Generic.List[string]
function Write-Log { param([string]$Message) $script:LogLines.Add($Message) }

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
function Get-Sources { param($Classification) ($Classification.Items | ForEach-Object { $_.Source }) -join ',' }
function New-File { param([string]$Name, $Sec) [pscustomobject]@{ Name = $Name; DurationSec = $Sec } }
function New-DdTitle {
    param([int]$Index, $Sec, [string]$Type = $null, $Season = $null, $Episode = $null, [string]$Name = "")
    [pscustomobject]@{ Index = $Index; SourceFile = ''; DurationSec = $Sec; Size = 0; ItemType = $Type; Season = $Season; Episode = $Episode; Name = $Name }
}
function New-DdDisc { param([object[]]$Titles) [pscustomobject]@{ Description = 'Test Show (2020) - Season 1, Disc 2'; Titles = $Titles } }

# A trimmed copy of a real TheDiscDB response (30 Rock, Season 1 Disc 2), round-tripped
# through JSON so the mock returns the same PSCustomObject shape Invoke-RestMethod does.
$hash302 = '16E974A41F04B04E0FC0F7B27EA29758'
function New-DiscDbTitleJson {
    param($Index, $Duration, $Type, $Season, $Episode, $Title)
    $item = if ($Type) { @{ title = $Title; type = $Type; season = "$Season"; episode = "$Episode" } } else { $null }
    @{ index = $Index; sourceFile = ('{0:D5}.mpls' -f $Index); duration = $Duration; size = 1000; itemType = $(if ($Type) { $Type } else { '' }); season = $(if ($Type) { "$Season" } else { '' }); episode = $(if ($Type) { "$Episode" } else { '' }); item = $item }
}
$matchResponseJson = @{
    data = @{ mediaItems = @{ nodes = @(
        @{ title = '30 Rock'; type = 'Series'; year = 2006; releases = @(
            @{ slug = 'the-complete-series-blu-ray'; title = 'The Complete Series'; discs = @(
                @{ index = 2; name = 'Season 1, Disc 2'; format = 'Blu-Ray'; contentHash = $hash302; titles = @(
                    (New-DiscDbTitleJson 0 '0:21:35' 'Episode' 1 8 'The Break-Up'),
                    (New-DiscDbTitleJson 1 '0:21:35' 'Episode' 1 9 'The Baby Show'),
                    (New-DiscDbTitleJson 2 '0:21:37' 'Episode' 1 10 'The Rural Juror'),
                    (New-DiscDbTitleJson 3 '0:00:09' $null 0 0 ''),
                    (New-DiscDbTitleJson 4 '0:04:52' 'DeletedScene' 1 10 'The Rural Juror: Deleted Scene?')
                ) }
            ) }
        ) }
    ) } }
} | ConvertTo-Json -Depth 12

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-discdb-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

function New-SizedFile { param([string]$Path, [int]$Bytes) New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null; [System.IO.File]::WriteAllBytes($Path, (New-Object byte[] $Bytes)) }

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nContent hash (TheDiscDB's algorithm)" -ForegroundColor Cyan

    Assert-Equal 'A4E9B5C484B45C3B7B3347F634A74108' (Get-DiscDbContentHashFromSizes -Sizes 5767010304, 23832576, 0, 1) 'MD5 over little-endian Int64 sizes, uppercase hex, no dashes (sizes over 4 GB included)'
    Assert-Equal '89AF2DF214317989ED233CBEFFE82F0C' (Get-DiscDbContentHashFromSizes -Sizes 12345) 'single-file vector'

    $bd = Join-Path $tempRoot 'bd'
    New-SizedFile (Join-Path $bd 'BDMV\STREAM\00002.m2ts') 5
    New-SizedFile (Join-Path $bd 'BDMV\STREAM\00000.m2ts') 10
    New-SizedFile (Join-Path $bd 'BDMV\STREAM\00001.m2ts') 20
    New-SizedFile (Join-Path $bd 'BDMV\STREAM\notes.txt') 7
    New-SizedFile (Join-Path $bd 'BDMV\STREAM\SSIF\00000.ssif') 99
    New-SizedFile (Join-Path $bd 'BDMV\index.bdmv') 3
    $h = Get-DiscDbContentHash -DriveRoot $bd
    Assert-Equal '33437DDFB33FCAD19AB50068E5C083CB' $h.Hash 'Blu-ray: only BDMV\STREAM\*.m2ts, sorted by name (not creation order), subfolders and other files ignored'
    Assert-Equal 'Blu-ray' $h.Format 'Blu-ray format reported'
    Assert-Equal 3 $h.FileCount 'three stream files counted'

    $dvd = Join-Path $tempRoot 'dvd'
    New-SizedFile (Join-Path $dvd 'VIDEO_TS\VTS_01_1.VOB') 8
    New-SizedFile (Join-Path $dvd 'VIDEO_TS\VIDEO_TS.IFO') 4
    New-SizedFile (Join-Path $dvd 'VIDEO_TS\VTS_01_0.IFO') 6
    New-SizedFile (Join-Path $dvd 'AUDIO_TS\ignored.dat') 50
    $h = Get-DiscDbContentHash -DriveRoot "$dvd\"
    Assert-Equal '0477A009AD3D1B459CB4C56B282CB9D5' $h.Hash 'DVD: every file directly in VIDEO_TS, sorted by name (trailing backslash on the root is fine)'
    Assert-Equal 'DVD' $h.Format 'DVD format reported'

    $empty = Join-Path $tempRoot 'empty'
    New-Item -ItemType Directory -Path $empty | Out-Null
    Assert-Equal $null (Get-DiscDbContentHash -DriveRoot $empty) 'no BDMV or VIDEO_TS -> $null'
    $odd = Join-Path $tempRoot 'odd'
    New-SizedFile (Join-Path $odd 'BDMV\index.bdmv') 3
    New-SizedFile (Join-Path $odd 'VIDEO_TS\VIDEO_TS.IFO') 4
    Assert-Equal $null (Get-DiscDbContentHash -DriveRoot $odd) 'a BDMV disc with no stream files does not fall back to VIDEO_TS (TheDiscDB does not either)'

    # ---------------------------------------------------------------------------
    Write-Host "`nLookup (mocked TheDiscDB)" -ForegroundColor Cyan

    $script:DiscDbBodies = New-Object System.Collections.Generic.List[string]
    function Invoke-DiscDbRequest {
        param([string]$Body)
        $script:DiscDbBodies.Add($Body)
        return ($matchResponseJson | ConvertFrom-Json)
    }

    $r = Get-SeriesDiscDbLookup -ContentHash $hash302.ToLower() 6>$null
    Assert-True ($null -ne $r.Disc) 'match: a known hash returns the disc (hash accepted in lower case)'
    Assert-Equal $hash302 $r.ContentHash 'match: hash normalised to upper case'
    Assert-Equal '30 Rock (2006) - The Complete Series, Season 1, Disc 2' $r.Disc.Description 'match: description names the show, release and disc'
    Assert-Equal 5 @($r.Disc.Titles).Count 'match: every listed title is returned, identified or not'
    Assert-Equal '8|Episode|1|1295' ("{0}|{1}|{2}|{3}" -f $r.Disc.Titles[0].Episode, $r.Disc.Titles[0].ItemType, $r.Disc.Titles[0].Season, $r.Disc.Titles[0].DurationSec) 'match: episode/season parsed to numbers, h:mm:ss to seconds'
    Assert-Equal $null $r.Disc.Titles[3].ItemType 'match: a title TheDiscDB has not identified has no item type'
    Assert-True ($r.Message -match '^TheDiscDB: matched 30 Rock' -and $r.Message -match '4 of 5 title') 'match: one-line message for the session log'
    $sent = $script:DiscDbBodies[0] | ConvertFrom-Json
    Assert-Equal $hash302 $sent.variables.hash 'request: the hash travels as a GraphQL variable, not pasted into the query'
    Assert-True ($sent.query -match 'contentHash: \{ eq: \$hash \}' -and $sent.query -match 'item \{ title type season episode \}') 'request: filters by contentHash and asks for each title''s item mapping'
    Assert-Equal 8 (Get-DiscDbFirstEpisode -DiscDbDisc $r.Disc -Season 1) 'first episode on the disc (default for the Disc 2+ prompt)'
    Assert-Equal $null (Get-DiscDbFirstEpisode -DiscDbDisc $r.Disc -Season 3) '...none for a season the disc does not hold'

    function Invoke-DiscDbRequest { param([string]$Body) return ('{"data":{"mediaItems":{"nodes":[]}}}' | ConvertFrom-Json) }
    $r = Get-SeriesDiscDbLookup -ContentHash 'ABCDEF0123456789ABCDEF0123456789' 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match 'not in the database') 'no match: Disc is $null with one explanatory line'
    Assert-Equal 'ABCDEF0123456789ABCDEF0123456789' $r.ContentHash 'no match: the hash is kept (so a later continue-rip.ps1 can retry)'

    function Invoke-DiscDbRequest { param([string]$Body) return ('{"errors":[{"message":"The field `bogus` does not exist"}]}' | ConvertFrom-Json) }
    $r = Get-SeriesDiscDbLookup -ContentHash $hash302 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match 'lookup failed' -and $r.Message -match 'bogus') 'GraphQL error: falls back, message carries the server error'

    function Invoke-DiscDbRequest { param([string]$Body) throw 'The operation has timed out.' }
    $threw = $false
    try { $r = Get-SeriesDiscDbLookup -ContentHash $hash302 6>$null } catch { $threw = $true }
    Assert-True (-not $threw -and $null -eq $r.Disc -and $r.Message -match 'timed out') 'offline / timeout: no exception, Disc is $null'

    function Invoke-DiscDbRequest { param([string]$Body) return 'not json at all' }
    $r = Get-SeriesDiscDbLookup -ContentHash $hash302 6>$null
    Assert-True ($null -eq $r.Disc) 'garbage response: treated as no match, no exception'

    function Invoke-DiscDbRequest { param([string]$Body) throw 'must not be called' }
    $r = Get-SeriesDiscDbLookup -ContentHash $hash302 -Disabled 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match '-NoDiscDb') '-NoDiscDb: no request at all'
    $r = Get-SeriesDiscDbLookup -ContentHash 'not-a-hash' 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match 'not a valid disc hash' -and $r.ContentHash -eq '') 'malformed -DiscDbHash: rejected without a request'
    $r = Get-SeriesDiscDbLookup -NoDriveReason 'no -DiscDbHash given' 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match 'no -DiscDbHash given') 'no hash and no drive: skipped with the caller''s reason'
    $r = Get-SeriesDiscDbLookup -DriveRoot $empty 6>$null
    Assert-True ($null -eq $r.Disc -and $r.Message -match 'no BDMV') 'drive without a video disc: skipped without a request'

    $script:DiscDbBodies.Clear()
    function Invoke-DiscDbRequest { param([string]$Body) $script:DiscDbBodies.Add($Body); return ('{"data":{"mediaItems":{"nodes":[]}}}' | ConvertFrom-Json) }
    $r = Get-SeriesDiscDbLookup -DriveRoot $bd 6>$null
    Assert-Equal '33437DDFB33FCAD19AB50068E5C083CB' (($script:DiscDbBodies[0] | ConvertFrom-Json).variables.hash) 'drive path: the hash read from the disc is what gets looked up'

    # ---------------------------------------------------------------------------
    Write-Host "`nLining files up with TheDiscDB titles" -ForegroundColor Cyan

    $disc = New-DdDisc @(
        (New-DdTitle 0 9),
        (New-DdTitle 1 1295 'Episode' 1 8 'A'),
        (New-DdTitle 2 1295 'Episode' 1 9 'B'),
        (New-DdTitle 3 1297 'Episode' 1 10 'C'),
        (New-DdTitle 4 2600),
        (New-DdTitle 5 292 'DeletedScene' 1 9 'B Deleted')
    )
    # The user's MakeMKV dropped the 9-second title 0, so every _tNN is one lower than
    # TheDiscDB's index - and B/C are within 2 seconds of A, so durations alone are
    # ambiguous. The play-all (index 4) was skipped by Step 2.
    $files = @((New-File 'x_t00.mp4' 1296), (New-File 'x_t01.mp4' 1294), (New-File 'x_t02.mp4' 1297), (New-File 'x_t04.mp4' 291))
    $map = Get-DiscDbTitleMatches -Titles $files -DiscDbDisc $disc
    Assert-Equal '1,2,3,5' (($files | ForEach-Object { $map[$_.Name].Index }) -join ',') 'index drift + near-identical durations: aligned in order (t00->1, t01->2, t02->3, t04->5)'

    $files = @((New-File 'title_t01.mp4' 1295), (New-File 'title_t02.mp4' 1295), (New-File 'title_t03.mp4' 1297))
    $map = Get-DiscDbTitleMatches -Titles $files -DiscDbDisc $disc
    Assert-Equal '1,2,3' (($files | ForEach-Object { $map[$_.Name].Index }) -join ',') 'no drift: same numbering lines up one to one'

    $files = @((New-File 'title_t00.mp4' 1295), (New-File 'title_t01.mp4' $null), (New-File 'title_t02.mp4' 1297))
    $map = Get-DiscDbTitleMatches -Titles $files -DiscDbDisc $disc
    Assert-True (-not $map.ContainsKey('title_t01.mp4') -and $map.ContainsKey('title_t00.mp4') -and $map.ContainsKey('title_t02.mp4')) 'a file with no readable duration is left unmatched; its neighbours still match'

    $files = @((New-File 'title_t00.mp4' 600), (New-File 'title_t01.mp4' 3600))
    Assert-Equal 0 (Get-DiscDbTitleMatches -Titles $files -DiscDbDisc $disc).Count 'durations that match nothing -> no pairs at all'
    Assert-Equal 0 (Get-DiscDbTitleMatches -Titles $files -DiscDbDisc $null).Count 'no disc record -> empty map'

    $files = @((New-File 'title_t00.mp4' 1310))
    Assert-Equal 0 (Get-DiscDbTitleMatches -Titles $files -DiscDbDisc (New-DdDisc @((New-DdTitle 0 1295 'Episode' 1 1 'A')))).Count 'a 15 s difference on a 21 min title is outside the 5 s / 1% tolerance'

    # ---------------------------------------------------------------------------
    Write-Host "`nClassification with TheDiscDB" -ForegroundColor Cyan

    $titles = @((New-File 'title_t00.mkv' 1295), (New-File 'title_t01.mkv' 1295), (New-File 'title_t02.mkv' 292), (New-File 'title_t03.mkv' 1300))
    $map = @{
        'title_t00.mkv' = (New-DdTitle 0 1295 'Episode' 1 8 'The Break-Up')
        'title_t01.mkv' = (New-DdTitle 1 1295 'Episode' 1 9 'The Baby Show')
        'title_t02.mkv' = (New-DdTitle 2 292 'DeletedScene' 1 9 'Making Of: "The Baby Show"?')
    }
    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 1 -DiscDbMap $map -Season 1
    Assert-Equal 'E8,E9,X1,E1' (Get-Kinds $c) 'TheDiscDB numbers used as published (E08, E09); the unmatched episode-length title numbered from -StartEpisode'
    Assert-Equal 'TheDiscDB,TheDiscDB,TheDiscDB,Length' (Get-Sources $c) 'Source column: TheDiscDB rows, then the median fallback for the rest'
    Assert-Equal 'Making Of The Baby Show' $c.Items[2].Label 'extra carries its published name, made filename-safe'
    Assert-Equal 'TheDiscDB: The Break-Up' $c.Items[0].Note 'episode note shows the published episode title'
    Assert-True ($c.Items[3].Mismatch -and $c.Items[3].Note -match 'not matched in TheDiscDB') 'an episode-length title TheDiscDB did not cover is flagged for checking'
    Assert-Equal 3 $c.DiscDbCount 'DiscDbCount counts the TheDiscDB-decided rows'
    Assert-Equal 1295 $c.Items[0].ExpectedSec 'Expected column shows TheDiscDB''s duration'

    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 8 -DiscDbMap $map -Season 1
    Assert-Equal 'E8,E9,X1,E10' (Get-Kinds $c) 'sequential numbering skips the numbers TheDiscDB reserved'

    $plan = New-SeriesRenamePlan -Classification $c -Title '30 Rock' -Season 1 -Directory $tempRoot
    Assert-Equal '30 Rock-S01-Extra01-Making Of The Baby Show.mkv' $plan[2].NewName 'named extra: <Title>-S##-Extra##-<published name>.ext'
    Assert-Equal 'TheDiscDB' $plan[0].Source 'plan rows carry the Source for the confirmation table'

    $c = Get-SeriesTitleClassification -Titles $titles -DiscDbMap $map -Season 1 -Overrides @{ 'title_t00.mkv' = 'Extra'; 'title_t02.mkv' = 'Episode' }
    Assert-Equal 'X1,E9,E1,E2' (Get-Kinds $c) 'an edit at the prompt beats TheDiscDB (and frees its number)'
    Assert-Equal 'You,TheDiscDB,You,Length' (Get-Sources $c) '...shown as Source You'

    $c = Get-SeriesTitleClassification -Titles $titles -DiscDbMap $map -Season 1 -AllExtras
    Assert-Equal 'X1,X2,X3,X4' (Get-Kinds $c) '-Extras disc beats TheDiscDB episodes'
    Assert-Equal 'Making Of The Baby Show' $c.Items[2].Label '...but a TheDiscDB extra keeps its name'

    $c = Get-SeriesTitleClassification -Titles $titles -DiscDbMap $map -Season 2
    Assert-True ($c.Items[0].Mismatch -and $c.Items[0].Note -match 'S01E08 - check' -and $c.Items[0].EpisodeNumber -eq 8) 'TheDiscDB season differs from -Season: number kept, row flagged'

    $dupMap = @{ 'title_t00.mkv' = (New-DdTitle 0 1295 'Episode' 1 8 'A'); 'title_t01.mkv' = (New-DdTitle 1 1295 'Episode' 1 8 'A again') }
    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 1 -DiscDbMap $dupMap -Season 1
    Assert-Equal 'E8,E1,X1,E2' (Get-Kinds $c) 'two files with the same TheDiscDB episode: the second falls back instead of reusing E08'

    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 1 -DiscDbMap $map -Season 1 -TakenEpisodes @(8)
    Assert-True ($c.Items[0].Source -ne 'TheDiscDB' -and $c.Items[0].EpisodeNumber -ne 8) 'a TheDiscDB number already used by a renamed file is not used twice'

    $unknownMap = @{ 'title_t00.mkv' = (New-DdTitle 0 1295) }
    $c = Get-SeriesTitleClassification -Titles $titles -DiscDbMap $unknownMap -Season 1
    Assert-True ($c.Items[0].Source -ne 'TheDiscDB') 'a title TheDiscDB lists but has not identified does not decide anything'

    # ---------------------------------------------------------------------------
    Write-Host "`nFallback order: TheDiscDB, then TMDb, then median" -ForegroundColor Cyan

    $tmdb = @(1..12 | ForEach-Object { [pscustomobject]@{ Number = $_; Name = "Ep $_"; Runtime = 22 } })
    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 10 -DiscDbMap $map -Season 1 -TmdbEpisodes $tmdb -TmdbEpisodeCount 12
    Assert-Equal 'TheDiscDB,TheDiscDB,TheDiscDB,TMDb' (Get-Sources $c) 'TheDiscDB first; TMDb runtimes for what TheDiscDB does not cover'
    Assert-Equal 'E8,E9,X1,E10' (Get-Kinds $c) '...TMDb-matched title numbered after the reserved ones'
    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 1 -TmdbEpisodes $tmdb -TmdbEpisodeCount 12
    Assert-Equal 'TMDb,TMDb,TMDb,TMDb' (Get-Sources $c) 'no TheDiscDB match: TMDb for everything (unchanged behaviour)'
    Assert-Equal 0 $c.DiscDbCount '...and no TheDiscDB rows'
    $c = Get-SeriesTitleClassification -Titles $titles -StartEpisode 1
    Assert-Equal 'Length,Length,Length,Length' (Get-Sources $c) 'neither: the median heuristic'
    Assert-Equal 'Median' $c.Method '...Method still reports Median'

    # ---------------------------------------------------------------------------
    Write-Host "`nNamed extras: label and pattern" -ForegroundColor Cyan

    Assert-Equal 'Making Of Season 2' (Get-SeriesExtraLabel 'Making Of: "Season 2"?') 'illegal filename characters removed, spaces collapsed'
    Assert-Equal 'Gag Reel' (Get-SeriesExtraLabel '  Gag Reel...  ') 'trailing dots/spaces trimmed (Windows would strip them silently)'
    Assert-Equal '' (Get-SeriesExtraLabel '???') 'nothing usable -> empty label (plain -Extra##)'
    $long = Get-SeriesExtraLabel 'An Extremely Long Behind The Scenes Featurette About Absolutely Everything'
    Assert-True ($long.Length -le 50 -and $long -notmatch ' $' -and 'An Extremely Long Behind The Scenes Featurette About Absolutely Everything'.StartsWith($long)) 'long names capped at 50 characters on a word boundary'
    Assert-Equal 'Show-S01-Extra03.mp4' (Get-SeriesExtraFileName -Title 'Show' -Season 1 -Extra 3 -Extension '.mp4' -Label '') 'no label -> unchanged -Extra## name'
    $pattern = Get-SeriesNamePattern -Title 'Show' -Season 1
    Assert-True ('Show-S01-Extra03-Making Of.mp4' -match $pattern -and $Matches['extra'] -eq '03') 'already-renamed pattern recognises a named extra (so a re-run leaves it alone)'
    Assert-True (-not ('Show-S01-E03-Making Of.mp4' -match $pattern)) '...but not an episode with a suffix'

    # ---------------------------------------------------------------------------
    Write-Host "`nRename end to end with a TheDiscDB match" -ForegroundColor Cyan

    $discDir = Join-Path $tempRoot 'Series\30 Rock\Season 1\Disc2'
    New-Item -ItemType Directory -Path $discDir -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4', 'title_t02.mp4', 'title_t03.mp4' | ForEach-Object { Set-Content -Path (Join-Path $discDir $_) -Value $_ }
    $durations = @{ 'title_t00.mp4' = 1295; 'title_t01.mp4' = 1296; 'title_t02.mp4' = 1297; 'title_t03.mp4' = 292 }
    $getDuration = { param($path) $durations[(Split-Path -Leaf $path)] }.GetNewClosure()
    function Invoke-DiscDbRequest { param([string]$Body) return ($matchResponseJson | ConvertFrom-Json) }
    $lookup = Get-SeriesDiscDbLookup -ContentHash $hash302 6>$null
    $script:LogLines.Clear()
    $result = Invoke-SeriesEpisodeRename -Directory $discDir -Title '30 Rock' -Season 1 -StartEpisode 8 -DiscDbDisc $lookup.Disc `
        -ReadInput (New-InputQueue @('')) -GetDuration $getDuration 6>$null
    $names = (Get-ChildItem -LiteralPath $discDir -File | Where-Object { $_.Extension -eq '.mp4' } | Sort-Object Name | ForEach-Object { $_.Name }) -join ','
    Assert-Equal '30 Rock-S01-E08.mp4,30 Rock-S01-E09.mp4,30 Rock-S01-E10.mp4,30 Rock-S01-Extra01-The Rural Juror Deleted Scene.mp4' $names 'episodes numbered by TheDiscDB; the deleted scene renamed with its published name'
    $rows = @(Import-Csv (Join-Path $discDir 'rename-manifest.csv'))
    Assert-Equal '30 Rock-S01-Extra01-The Rural Juror Deleted Scene.mp4' $rows[3].NewName 'manifest records the named extra (undo-rename.ps1 works from it unchanged)'
    Assert-True (@($script:LogLines | Where-Object { $_ -match '^TheDiscDB: 4 of 4 file\(s\) lined up' }).Count -eq 1) 'how many files lined up is logged'
    Assert-True (@($script:LogLines | Where-Object { $_ -match 'classification by TheDiscDB \(4 title' }).Count -eq 1) 'the session log records that TheDiscDB decided the names'
    $again = Invoke-SeriesEpisodeRename -Directory $discDir -Title '30 Rock' -Season 1 -DiscDbDisc $lookup.Disc -ReadInput (New-InputQueue @('')) -GetDuration $getDuration 6>$null
    Assert-Equal 0 $again.Renamed 're-running organize leaves the named extra alone'

    $discDir2 = Join-Path $tempRoot 'Series\30 Rock\Season 1\Disc9'
    New-Item -ItemType Directory -Path $discDir2 -Force | Out-Null
    'title_t00.mp4', 'title_t01.mp4' | ForEach-Object { Set-Content -Path (Join-Path $discDir2 $_) -Value $_ }
    $durations2 = @{ 'title_t00.mp4' = 2700; 'title_t01.mp4' = 2650 }
    $getDuration2 = { param($path) $durations2[(Split-Path -Leaf $path)] }.GetNewClosure()
    $result = Invoke-SeriesEpisodeRename -Directory $discDir2 -Title '30 Rock' -Season 1 -StartEpisode 1 -DiscDbDisc $lookup.Disc `
        -ReadInput (New-InputQueue @('')) -GetDuration $getDuration2 6>$null
    $names = (Get-ChildItem -LiteralPath $discDir2 -File | Where-Object { $_.Extension -eq '.mp4' } | Sort-Object Name | ForEach-Object { $_.Name }) -join ','
    Assert-Equal '30 Rock-S01-E01.mp4,30 Rock-S01-E02.mp4' $names 'matched disc but no file lines up: silent fallback to the median heuristic, rip not blocked'

    # ---------------------------------------------------------------------------
    Write-Host "`nOpt-out and hand-off to continue-rip.ps1" -ForegroundColor Cyan

    Assert-True ((Get-ScriptParameterNames $ripDiscPath) -contains 'NoDiscDb') 'rip-disc.ps1 has a -NoDiscDb switch'
    $continueParams = Get-ScriptParameterNames $continuePath
    Assert-True ($continueParams -contains 'NoDiscDb' -and $continueParams -contains 'DiscDbHash') 'continue-rip.ps1 accepts -NoDiscDb and -DiscDbHash'

    $cmd = Get-ContinueRipCommand -Title 'Fargo' -RemainingStepNumbers @(3) -Series -Season 1 -Disc 2 -StartEpisode 5 -EpisodeNames @() -DiscDbHash $hash302
    Assert-Equal ".\continue-rip.ps1 -title `"Fargo`" -FromStep organize -Series -Season 1 -Disc 2 -StartEpisode 5 -DiscDbHash $hash302" $cmd 'retry command carries the disc hash so continue-rip.ps1 can repeat the lookup'
    $cmd = Get-ContinueRipCommand -Title 'Fargo' -RemainingStepNumbers @(3) -Series -Season 1 -Disc 1 -StartEpisode 1 -EpisodeNames @() -DiscDbHash $hash302 -NoDiscDb
    Assert-Equal '.\continue-rip.ps1 -title "Fargo" -FromStep organize -Series -Season 1 -NoDiscDb' $cmd '-NoDiscDb is carried over (and the hash dropped)'
    $cmd = Get-ContinueRipCommand -Title 'Inception' -RemainingStepNumbers @(2) -Disc 1 -StartEpisode 1 -EpisodeNames @() -DiscDbHash $hash302
    Assert-Equal '.\continue-rip.ps1 -title "Inception" -FromStep handbrake' $cmd 'no hash in a movie command (TheDiscDB is only used for series naming)'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
