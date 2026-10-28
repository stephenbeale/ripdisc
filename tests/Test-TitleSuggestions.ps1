#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for inferring the title from the disc: title keys, generic-label detection,
    the non-interactive TMDb candidate search, merging TheDiscDB / disc label / TMDb into
    one menu (with years and IMDb ids), the IMDb suggestion source, the numbered pick prompt, and the -title vs disc-label sanity check.

.DESCRIPTION
    Lifts the real functions out of rip-disc.ps1 with the AST parser and dot-sources the
    real SeriesEpisodes.ps1, so nothing here is a reimplementation. TMDb and IMDb are mocked by
    replacing the Invoke-TMDbRequest / Invoke-ImdbSuggestRequest seams - no network. No disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-TitleSuggestions.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$sharedPath = Join-Path $repoRoot 'SeriesEpisodes.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

foreach ($p in @($sharedPath, $ripDiscPath)) {
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
foreach ($name in @('Clean-DiscName', 'New-TitleCandidate', 'Get-ShortText', 'Get-TitleCandidateUrl', 'Show-TitleCandidates', 'ConvertTo-TitleKey', 'Test-GenericDiscLabel', 'Get-TMDbTitleCandidates', 'Get-ImdbTitleCandidates', 'Format-TitleCandidate',
                    'Get-DiscTitleCandidates', 'Read-TitleChoice', 'Get-DiscLabelTitleCheck', 'Get-RipAudioCommand', 'Get-OtherDiscDrives', 'Read-OtherDrive', 'Get-RelaunchParams')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $ripDiscPath $name)))
}

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
        Write-Host "        expected: $Expected" -ForegroundColor Red
        Write-Host "        actual:   $Actual" -ForegroundColor Red
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

function Format-Candidates { param($List) (@($List) | ForEach-Object { "$($_.Title)|$($_.Year)|$($_.MediaType)|$($_.ImdbId)|$($_.Source)" }) -join ';' }

# TMDb mock: records the URL, returns a JSON round-tripped body like Invoke-RestMethod.
$script:TmdbUrls = @()
$script:TmdbBody = $null
$script:TmdbThrow = $false
function Invoke-TMDbRequest {
    param([string]$Url)
    $script:TmdbUrls += $Url
    if ($script:TmdbThrow) { throw "network down" }
    if ($Url -match '/(movie|tv)/(\d+)/external_ids') {
        if ($Matches[2] -eq '303') { throw "404" }
        if ($Matches[2] -eq '404') { return [pscustomobject]@{ imdb_id = $null } }
        return [pscustomobject]@{ imdb_id = "tt$($Matches[2])000" }
    }
    return $script:TmdbBody
}
$script:ImdbUrls = @()
$script:ImdbBody = $null
$script:ImdbThrow = $false
function Invoke-ImdbSuggestRequest {
    param([string]$Url)
    $script:ImdbUrls += $Url
    if ($script:ImdbThrow) { throw "network down" }
    return $script:ImdbBody
}
# Shape of a real response (fetched 2026-10-28 for "30 rock"), trimmed, plus one of
# each kind that must be dropped.
$rockImdbJson = @{ d = @(
    @{ id = 'tt0496424'; l = '30 Rock'; y = 2006; q = 'TV series'; qid = 'tvSeries'; s = 'Tina Fey, Alec Baldwin' },
    @{ id = 'tt1860375'; l = '30 Rock'; y = 1983; q = 'TV series'; qid = 'tvSeries' },
    @{ id = 'tt7306384'; l = 'Keri Hilson: 30 Rock'; y = 2009; q = 'musicVideo'; qid = 'musicVideo' },
    @{ id = 'nm0000001'; l = 'Some Person'; s = 'Actor' },
    @{ id = 'tt0052529'; l = 'The 30 Foot Bride of Candy Rock'; y = 1959; q = 'feature'; qid = 'movie' },
    @{ id = 'tt2866584'; l = '30 Rock: The Webisodes'; q = 'video'; qid = 'video' }
) } | ConvertTo-Json -Depth 5
$slowDanceJson = @{ results = @(
    @{ media_type = 'movie'; title = 'Slow Dance'; release_date = '2013-05-01'; id = 101; overview = 'Two surfers chase a winter swell along the coast of Portugal, and the film follows them for weeks through storms and flat spells alike.' },
    @{ media_type = 'person'; name = 'Slow Dancer' ; id = 7 },
    @{ media_type = 'tv'; name = 'Slow Dance Club'; first_air_date = '2019-02-02'; id = 202 },
    @{ media_type = 'movie'; title = 'Slow Dance'; release_date = '1998-01-01'; id = 303 },
    @{ media_type = 'movie'; title = 'No Date'; release_date = $null; id = 404 }
) } | ConvertTo-Json -Depth 5

# ---------------------------------------------------------------------------
Write-Host "`nTitle keys and generic labels" -ForegroundColor Cyan

Assert-Equal (ConvertTo-TitleKey 'Slow Dance') (ConvertTo-TitleKey 'SLOW_DANCE') 'case and underscores ignored'
Assert-Equal (ConvertTo-TitleKey 'Slow Dance') (ConvertTo-TitleKey 'The Slow Dance') 'a leading "The" ignored'
Assert-Equal (ConvertTo-TitleKey 'Fast & Loose') (ConvertTo-TitleKey 'Fast and Loose') '"&" equals "and"'
Assert-Equal 'theatre' (ConvertTo-TitleKey 'Theatre') '"The" inside a word is kept'
foreach ($g in @('Audio CD', 'AUDIO_CD', 'DVD_VIDEO', 'DVDVIDEO', 'Disc1', 'BDROM', 'LOGICAL_VOLUME_ID', '12345', '', '  ', 'VIDEO_TS')) {
    Assert-True (Test-GenericDiscLabel $g) "'$g' is a generic label"
}
foreach ($g in @('Slow Dance', 'SILICON_VALLEY_S1_D2', 'Disco Inferno')) {
    Assert-True (-not (Test-GenericDiscLabel $g)) "'$g' is a real label"
}

# ---------------------------------------------------------------------------
Write-Host "`nTMDb candidates (mocked)" -ForegroundColor Cyan

$savedKey = $env:TMDB_API_KEY
$script:Config_TmdbApiKey = $null
try {
    $env:TMDB_API_KEY = $null
    $c = Get-TMDbTitleCandidates -SearchTitle 'Slow Dance' 6>$null
    Assert-Equal 0 @($c).Count 'no key: empty list, no request'
    Assert-Equal 0 $script:TmdbUrls.Count 'no key: TMDb not called'

    $env:TMDB_API_KEY = 'k'
    $script:TmdbBody = $slowDanceJson | ConvertFrom-Json
    $c = Get-TMDbTitleCandidates -SearchTitle 'Slow Dance' 6>$null
    Assert-Equal 'Slow Dance|2013|movie|tt101000|TMDb;Slow Dance Club|2019|tv|tt202000|TMDb;Slow Dance|1998|movie||TMDb;No Date||movie||TMDb' (Format-Candidates $c) 'movies and TV kept in order, people dropped, missing date gives no year; IMDb ids fetched, a failed or empty lookup leaves none'
    Assert-True ($script:TmdbUrls[0] -match 'search/multi\?query=Slow%20Dance&api_key=k$') 'title URL-encoded into a multi search'
    Assert-True ($script:TmdbUrls[1] -match '/movie/101/external_ids\?api_key=k$' -and $script:TmdbUrls[2] -match '/tv/202/external_ids') 'external_ids asked per candidate with the right media type'
    Assert-Equal 101 $c[0].TmdbId 'TMDb id kept for later season lookups'
    Assert-Equal 'Two surfers chase a winter swell along the coast of Portugal, and the film follows them...' $c[0].Description 'TMDb overview shortened at a word boundary'
    Assert-Equal 'Movie|TV' "$($c[0].Kind)|$($c[1].Kind)" 'TMDb kind shown as Movie / TV'
    $c = Get-TMDbTitleCandidates -SearchTitle 'Slow Dance' -Max 2 6>$null
    Assert-Equal 2 @($c).Count '-Max caps the list'

    $script:TmdbThrow = $true
    $c = Get-TMDbTitleCandidates -SearchTitle 'Slow Dance' 6>$null
    Assert-Equal 0 @($c).Count 'request failure: empty list, never throws'
    $script:TmdbThrow = $false
} finally {
    $env:TMDB_API_KEY = $savedKey
}

# ---------------------------------------------------------------------------
Write-Host "`nIMDb candidates (mocked)" -ForegroundColor Cyan

$script:ImdbBody = $rockImdbJson | ConvertFrom-Json
$c = Get-ImdbTitleCandidates -SearchTitle '30 Rock' 6>$null
Assert-Equal '30 Rock|2006|tv|tt0496424|IMDb;30 Rock|1983|tv|tt1860375|IMDb;The 30 Foot Bride of Candy Rock|1959|movie|tt0052529|IMDb;30 Rock: The Webisodes||movie|tt2866584|IMDb' (Format-Candidates $c) 'titles kept with id/year/kind; music videos and people dropped; missing year left blank'
Assert-Equal 'https://v2.sg.media-imdb.com/suggestion/3/30%20rock.json' $script:ImdbUrls[-1] 'URL: lowercased query under its first character'
Assert-Equal 'with Tina Fey, Alec Baldwin' $c[0].Description 'IMDb cast becomes the description'
Assert-Equal 'TV series|Movie|Video' "$($c[0].Kind)|$($c[2].Kind)|$($c[3].Kind)" 'IMDb kind kept (TV series / Movie / Video), not lumped into Movie'
Assert-Equal '30 Rock (2006) [TV series]' (Format-TitleCandidate $c[0]) 'menu line uses the IMDb kind'
$null = Get-ImdbTitleCandidates -SearchTitle '"Weird" Al' 6>$null
Assert-True ($script:ImdbUrls[-1] -match '/suggestion/x/') 'a query not starting with a letter/digit goes under x'
$c = Get-ImdbTitleCandidates -SearchTitle '30 Rock' -Max 1 6>$null
Assert-Equal 1 @($c).Count '-Max caps the list'
$n = $script:ImdbUrls.Count
$c = Get-ImdbTitleCandidates -SearchTitle '   ' 6>$null
Assert-True (@($c).Count -eq 0 -and $script:ImdbUrls.Count -eq $n) 'blank search: no request'
$script:ImdbThrow = $true
$c = Get-ImdbTitleCandidates -SearchTitle '30 Rock' 6>$null
Assert-Equal 0 @($c).Count 'unofficial endpoint down: empty list, never throws'
$script:ImdbThrow = $false
$script:ImdbBody = [pscustomobject]@{ unexpected = 1 }
$c = Get-ImdbTitleCandidates -SearchTitle '30 Rock' 6>$null
Assert-Equal 0 @($c).Count 'changed response shape: empty list'

# ---------------------------------------------------------------------------
Write-Host "`nMerging TheDiscDB, disc label, TMDb and IMDb" -ForegroundColor Cyan

$tmdb = @(
    [pscustomobject]@{ Title = 'Slow Dance'; Year = '2013'; MediaType = 'movie'; TmdbId = 101; ImdbId = 'tt101000'; Season = 0; Source = 'TMDb' },
    [pscustomobject]@{ Title = 'Slow Dance Club'; Year = '2019'; MediaType = 'tv'; TmdbId = 202; ImdbId = $null; Season = 0; Source = 'TMDb' },
    [pscustomobject]@{ Title = 'Slow Dance'; Year = '1998'; MediaType = 'movie'; TmdbId = 303; ImdbId = $null; Season = 0; Source = 'TMDb' }
)
$c = Get-DiscTitleCandidates -DiscLabel 'SLOW_DANCE' -TmdbCandidates $tmdb
Assert-Equal 'Slow Dance||||disc label;Slow Dance|2013|movie|tt101000|TMDb;Slow Dance Club|2019|tv||TMDb;Slow Dance|1998|movie||TMDb' (Format-Candidates $c) 'the label stays its own line; same-name search hits are listed separately, never merged into it'
Assert-Equal 101 $c[1].TmdbId 'TMDb entry keeps its TMDb id'

$c = Get-DiscTitleCandidates -DiscLabel 'SLOW_DANCE'
Assert-Equal 'Slow Dance||||disc label' (Format-Candidates $c) 'no TMDb key: the cleaned label is still offered'

$c = Get-DiscTitleCandidates -DiscLabel 'DVD_VIDEO' -TmdbCandidates @()
Assert-Equal 0 @($c).Count 'generic label and nothing else: no candidates'

$dd = [pscustomobject]@{ MediaTitle = '30 Rock'; MediaType = 'Series'; Year = 2006; Titles = @(
    [pscustomobject]@{ ItemType = 'Episode'; Season = 1 },
    [pscustomobject]@{ ItemType = 'Episode'; Season = 1 },
    [pscustomobject]@{ ItemType = 'DeletedScene'; Season = 3 },
    [pscustomobject]@{ ItemType = $null; Season = $null }) }
$c = Get-DiscTitleCandidates -DiscLabel '30ROCK_S1_D2' -DiscDbDisc $dd -TmdbCandidates @(
    [pscustomobject]@{ Title = '30 Rock'; Year = '2006'; MediaType = 'tv'; TmdbId = 4608; ImdbId = $null; Season = 0; Source = 'TMDb' }) -ImdbCandidates @(
    [pscustomobject]@{ Title = '30 Rock'; Year = '2006'; MediaType = 'tv'; TmdbId = $null; ImdbId = 'tt0496424'; Season = 0; Source = 'IMDb' },
    [pscustomobject]@{ Title = '30 Rock'; Year = '1983'; MediaType = 'tv'; TmdbId = $null; ImdbId = 'tt1860375'; Season = 0; Source = 'IMDb' })
Assert-Equal '30 Rock|2006|tv|tt0496424|TheDiscDB + TMDb + IMDb;30Rock S1||||disc label;30 Rock|1983|tv|tt1860375|IMDb' (Format-Candidates $c) 'TheDiscDB first, TMDb and same-year IMDb fold into it; the 1983 show stays separate; an unsplit label (Clean-DiscName keeps S1) is its own option'
Assert-Equal 1 $c[0].Season 'season taken from TheDiscDB episodes (extras ignored)'
Assert-Equal 4608 $c[0].TmdbId 'TheDiscDB entry borrows the TMDb id'

$dd2 = [pscustomobject]@{ MediaTitle = 'Heat'; MediaType = 'Movie'; Year = 1995; Titles = @() }
$c = Get-DiscTitleCandidates -DiscLabel 'HEAT' -DiscDbDisc $dd2
Assert-Equal 'Heat|1995|movie||TheDiscDB + disc label' (Format-Candidates $c) 'movie from TheDiscDB, label agrees'
Assert-Equal 0 $c[0].Season 'no season for a movie'

$c = Get-DiscTitleCandidates -DiscLabel 'HEAT' -ImdbCandidates @(
    [pscustomobject]@{ Title = 'Heat'; Year = '1995'; MediaType = 'movie'; TmdbId = $null; ImdbId = 'tt0113277'; Season = 0; Source = 'IMDb' },
    [pscustomobject]@{ Title = 'Heat'; Year = '1986'; MediaType = 'movie'; TmdbId = $null; ImdbId = 'tt0093164'; Season = 0; Source = 'IMDb' })
Assert-Equal 'Heat||||disc label;Heat|1995|movie|tt0113277|IMDb;Heat|1986|movie|tt0093164|IMDb' (Format-Candidates $c) 'label first, then each same-name film on its own line'
Assert-Equal 'Heat (1995)' (Format-TitleCandidate $c[1]) 'menu line: title and year only; no [Movie], no IMDb id or source (the link line has them)'
Assert-Equal 'Slow Dance (2014) [Short]' (Format-TitleCandidate (New-TitleCandidate -Title 'Slow Dance' -Year '2014' -MediaType 'movie' -Kind 'Short' -ImdbId 'tt3453042' -Source 'IMDb')) 'other kinds are still flagged'
Assert-Equal 'Heat (1995) - exact disc match' (Format-TitleCandidate (New-TitleCandidate -Title 'Heat' -Year '1995' -MediaType 'movie' -Kind 'Movie' -Source 'TheDiscDB + IMDb')) 'a TheDiscDB row says it is an exact disc match'
Assert-Equal 'Heat - typed' (Format-TitleCandidate (New-TitleCandidate -Title 'Heat' -Source 'typed')) 'a typed title says so'
Assert-Equal 'Heat (1995) [IMDb tt0113277, TMDb 949, via TMDb + IMDb]' (Format-TitleCandidate (New-TitleCandidate -Title 'Heat' -Year '1995' -Kind 'Movie' -ImdbId 'tt0113277' -TmdbId 949 -Source 'TMDb + IMDb') -WithIds) '-WithIds (session log) keeps the ids and sources'
Assert-Equal 'Heat - disc label [via disc label]' (Format-TitleCandidate (New-TitleCandidate -Title 'Heat' -Source 'disc label') -WithIds) '-WithIds with no ids'
Assert-Equal 'Heat - disc label' (Format-TitleCandidate ([pscustomobject]@{ Title = 'Heat'; Year = ''; MediaType = $null; ImdbId = $null; Source = 'disc label' })) 'menu line with nothing extra known'

# Regression (live run 2026-10-28): label "Slow Dance" on a surf film took the year,
# [TV] type and IMDb id of IMDb's first hit, a 2005 Japanese TV series - Enter would
# then have switched Series mode on.
$c = Get-DiscTitleCandidates -DiscLabel 'Slow Dance' -ImdbCandidates @(
    [pscustomobject]@{ Title = 'Slow Dance'; Year = '2005'; MediaType = 'tv'; TmdbId = $null; ImdbId = 'tt0481050'; Season = 0; Source = 'IMDb' },
    [pscustomobject]@{ Title = 'Slow Dance'; Year = '2020'; MediaType = 'movie'; TmdbId = $null; ImdbId = 'tt12892860'; Season = 0; Source = 'IMDb' })
Assert-Equal 'Slow Dance||||disc label' (Format-Candidates @($c)[0]) 'the default pick is the plain label: no borrowed year, type or id'
Assert-Equal 3 @($c).Count 'both IMDb hits still offered'
$c = Get-DiscTitleCandidates -TmdbCandidates @(
    [pscustomobject]@{ Title = 'Heat'; Year = '1995'; MediaType = 'movie'; TmdbId = 949; ImdbId = 'tt0113277'; Season = 0; Source = 'TMDb' }) -ImdbCandidates @(
    [pscustomobject]@{ Title = 'Heat'; Year = '1995'; MediaType = 'movie'; TmdbId = $null; ImdbId = 'tt0113277'; Season = 0; Source = 'IMDb' },
    [pscustomobject]@{ Title = 'Heat'; Year = '1995'; MediaType = 'movie'; TmdbId = $null; ImdbId = 'tt9999999'; Season = 0; Source = 'IMDb' })
Assert-Equal 'Heat|1995|movie|tt0113277|TMDb + IMDb;Heat|1995|movie|tt9999999|IMDb' (Format-Candidates $c) 'TMDb and IMDb agreeing on the id merge; a different id stays separate even with the same year'

$c = Get-DiscTitleCandidates -DiscLabel '1917'
Assert-Equal 0 @($c).Count 'a bare-number label is treated as generic (known trade-off: the film 1917)'

# ---------------------------------------------------------------------------
Write-Host "`nPick prompt" -ForegroundColor Cyan

$list = Get-DiscTitleCandidates -DiscLabel 'SLOW_DANCE' -TmdbCandidates $tmdb
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('')) 6>$null
Assert-Equal 'Slow Dance|disc label' "$($pick.Title)|$($pick.Source)" 'Enter takes the first suggestion (the plain label)'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('4')) 6>$null
Assert-Equal 'Slow Dance|1998' "$($pick.Title)|$($pick.Year)" 'a number picks that row'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('9', 'x', '3')) 6>$null
Assert-Equal 'Slow Dance Club' $pick.Title 'out-of-range and junk answers re-ask'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('m', '  Slow Dance (Director''s Cut) ')) 6>$null
Assert-Equal "Slow Dance (Director's Cut)|typed" "$($pick.Title)|$($pick.Source)" 'M types a title (trimmed)'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('M', '')) 6>$null
Assert-True ($null -eq $pick) 'M then blank aborts'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('a')) 6>$null
Assert-True ($null -eq $pick) 'A aborts'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @()) 6>$null
Assert-True ($null -eq $pick) 'end of input aborts (never silently accepts)'
$pick = Read-TitleChoice -Candidates @() -ReadInput (New-InputQueue @('Slow Dance')) 6>$null
Assert-Equal 'Slow Dance|typed' "$($pick.Title)|$($pick.Source)" 'no suggestions: asks for the title straight away'
$pick = Read-TitleChoice -Candidates @() -ReadInput (New-InputQueue @('')) 6>$null
Assert-True ($null -eq $pick) 'no suggestions and blank: abort'

# ---------------------------------------------------------------------------
Write-Host "`nDescriptions, links and search again" -ForegroundColor Cyan

Assert-Equal 'https://www.imdb.com/title/tt0113277/' (Get-TitleCandidateUrl (New-TitleCandidate -Title 'Heat' -ImdbId 'tt0113277' -TmdbId 949 -Source 'IMDb')) 'IMDb page when the id is known'
Assert-Equal 'https://www.themoviedb.org/tv/4608' (Get-TitleCandidateUrl (New-TitleCandidate -Title '30 Rock' -MediaType 'tv' -TmdbId 4608 -Source 'TMDb')) 'TMDb page when only the TMDb id is known'
Assert-Equal 'https://www.imdb.com/find/?q=Slow%20Dance&s=tt' (Get-TitleCandidateUrl (New-TitleCandidate -Title 'Slow Dance' -Source 'disc label')) 'disc label row links to an IMDb title search'
Assert-Equal 'short text' (Get-ShortText '  short   text ') 'short text only tidied'
$c = Get-DiscTitleCandidates -DiscLabel 'SLOW_DANCE'
Assert-Equal 'as printed on the disc' $c[0].Description 'label row says where it came from'
$dd3 = [pscustomobject]@{ MediaTitle = 'Heat'; MediaType = 'Movie'; Year = 1995; Description = 'Heat (1995) - Blu-ray, disc 1'; Titles = @() }
$c = Get-DiscTitleCandidates -DiscDbDisc $dd3 -ImdbCandidates @(New-TitleCandidate -Title 'Heat' -Year '1995' -MediaType 'movie' -Kind 'Movie' -Description 'with Al Pacino' -ImdbId 'tt0113277' -Source 'IMDb')
Assert-Equal 'this exact disc: Heat (1995) - Blu-ray, disc 1|tt0113277' "$($c[0].Description)|$($c[0].ImdbId)" 'TheDiscDB keeps its own description and borrows the IMDb id'

$script:SearchedFor = @()
$searchStub = { param($q) $script:SearchedFor += $q; @(New-TitleCandidate -Title 'Slow Dance' -Year '2014' -MediaType 'movie' -Kind 'Short' -ImdbId 'tt3453042' -Source 'IMDb') }
$first = @(New-TitleCandidate -Title 'Slow Dance' -Source 'disc label')
$pick = Read-TitleChoice -Candidates $first -Search $searchStub -ReadInput (New-InputQueue @('r', 'slow dance 2014', '1')) 6>$null
Assert-Equal 'slow dance 2014' ($script:SearchedFor -join ';') 'R runs the search with the typed text'
Assert-Equal 'Slow Dance|2014|tt3453042' "$($pick.Title)|$($pick.Year)|$($pick.ImdbId)" 'the new list replaces the old one and can be picked from'
$pick = Read-TitleChoice -Candidates $first -Search $searchStub -ReadInput (New-InputQueue @('r', '', '')) 6>$null
Assert-Equal 'disc label' $pick.Source 'R then blank keeps the old list'
$pick = Read-TitleChoice -Candidates @() -Search { param($q) @() } -ReadInput (New-InputQueue @('r', 'nothing', 'a')) 6>$null
Assert-True ($null -eq $pick) 'no suggestions with search: R/M/A offered instead of jumping to typing'
$pick = Read-TitleChoice -Candidates $first -ReadInput (New-InputQueue @('r', '1')) 6>$null
Assert-Equal 'disc label' $pick.Source 'R without a search is just an invalid choice'

# ---------------------------------------------------------------------------
Write-Host "`n-title vs disc label" -ForegroundColor Cyan

$r = Get-DiscLabelTitleCheck -Title 'Slow Dance' -DiscLabel 'Slow Dance'
Assert-True ($r.Agrees) 'the reported example: label "Slow Dance" agrees with -title "Slow Dance"'
$r = Get-DiscLabelTitleCheck -Title 'Silicon Valley' -DiscLabel 'SILICON VALLEY S1 D2'
Assert-True ($r.Agrees) 'season/disc tokens in the label still agree'
$r = Get-DiscLabelTitleCheck -Title 'Point Break' -DiscLabel 'SLOW_DANCE'
Assert-True (-not $r.Agrees) 'a different label disagrees'
Assert-Equal 'Slow Dance' $r.Label 'the cleaned label is reported'
Assert-True ($null -eq (Get-DiscLabelTitleCheck -Title 'Anything' -DiscLabel 'DVD_VIDEO')) 'generic label: no check'
Assert-True ($null -eq (Get-DiscLabelTitleCheck -Title 'Anything' -DiscLabel '')) 'no label: no check'
$r = Get-DiscLabelTitleCheck -Title 'It' -DiscLabel 'SPLIT_DECISION'
Assert-True (-not $r.Agrees) 'very short titles do not match by containment'

# ---------------------------------------------------------------------------
Write-Host "`nAudio CD guard" -ForegroundColor Cyan

Assert-Equal '.\rip-audio.ps1 -Drive D' (Get-RipAudioCommand -DriveLetter 'D:') 'plain command, drive colon dropped'
Assert-Equal '.\rip-audio.ps1 -Drive D -OutputDrive F' (Get-RipAudioCommand -DriveLetter 'd:' -OutputDrive 'F:\') 'output drive carried when given; drive letters upper-cased'
Assert-Equal 'cd "C:\repos\ripaudio"; .\rip-audio.ps1 -Drive D' (Get-RipAudioCommand -DriveLetter 'D:' -RipAudioDir 'C:\repos\ripaudio') 'cd into the sibling ripaudio checkout when found'
$ripText = Get-Content -LiteralPath $ripDiscPath -Raw
$guardAt = $ripText.IndexOf('# ========== AUDIO CD GUARD ==========')
$discoveryAt = $ripText.IndexOf('AUTO-DISCOVERY MODE')
Assert-True ($guardAt -gt 0 -and $guardAt -lt $discoveryAt) 'the audio CD guard runs before auto-discovery (the reported run offered film titles for "Audio CD")'
Assert-True ($ripText.Substring($guardAt, 600) -match "Get-DiscTypeLabel -DriveLetter \`$driveLetter\) -eq 'Audio CD'") 'guard uses the same disc-type check as the drive listing'

# ---------------------------------------------------------------------------
Write-Host "`nUse a different drive" -ForegroundColor Cyan

# The reported run: D: held an audio CD, G: held SCARFACE.
$drv = @(
    [pscustomobject]@{ Index = 0; Letter = 'D:'; DiscName = 'Audio CD'; Busy = $false },
    [pscustomobject]@{ Index = 1; Letter = 'G:'; DiscName = 'SCARFACE'; Busy = $false },
    [pscustomobject]@{ Index = 2; Letter = 'H:'; DiscName = 'BUSY ONE'; Busy = $true },
    [pscustomobject]@{ Index = 3; Letter = 'I:'; DiscName = ''; Busy = $false },
    [pscustomobject]@{ Index = 4; Letter = ''; DiscName = 'NO LETTER'; Busy = $false })
$types = @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video'; 'H:' = 'DVD-Video'; 'I:' = '' }
$getType = { param($l) $types[$l] }.GetNewClosure()
$others = Get-OtherDiscDrives -DrvLines $drv -CurrentLetter 'd:' -GetDiscType $getType
Assert-Equal 'G:|SCARFACE|DVD-Video' (($others | ForEach-Object { "$($_.Letter)|$($_.DiscName)|$($_.DiscType)" }) -join ';') 'only G: offered: current drive, busy drive, empty drive and letterless entry left out'
$others = Get-OtherDiscDrives -DrvLines $drv -CurrentLetter 'G:' -GetDiscType $getType
Assert-Equal 'D:' (($others | ForEach-Object { $_.Letter }) -join ';') 'from G:, the audio CD drive is offered (the title list keeps every disc type)'
Assert-Equal 0 @(Get-OtherDiscDrives -DrvLines @() -CurrentLetter 'D:').Count 'no drive list (-DriveIndex mode): nothing offered'

$two = @([pscustomobject]@{ Letter = 'G:'; DiscName = 'SCARFACE'; DiscType = 'DVD-Video' }, [pscustomobject]@{ Letter = 'E:'; DiscName = ''; DiscType = 'Blu-ray' })
Assert-Equal 'E:' (Read-OtherDrive -Drives $two -ReadInput (New-InputQueue @('2')) 6>$null) 'pick by number'
Assert-Equal 'G:' (Read-OtherDrive -Drives $two -ReadInput (New-InputQueue @('x', 'g')) 6>$null) 'pick by letter after an invalid answer'
Assert-True ($null -eq (Read-OtherDrive -Drives $two -ReadInput (New-InputQueue @('')) 6>$null)) 'Enter goes back'
Assert-True ($null -eq (Read-OtherDrive -Drives $two -ReadInput (New-InputQueue @()) 6>$null)) 'end of input goes back'

$bound = @{ Drive = 'd'; DriveIndex = 0; OutputDrive = 'F'; Surf = [switch]$true; title = 'Slow Dance' }
$re = Get-RelaunchParams -BoundParams $bound -DriveLetter 'G:'
Assert-Equal 'Drive=G:;OutputDrive=F;Surf=True;title=Slow Dance' (($re.Keys | Sort-Object | ForEach-Object { "$_=$($re[$_])" }) -join ';') 'restart keeps every option, drops -DriveIndex, points -Drive at the new drive'
Assert-Equal 'd' $bound.Drive 'the original parameters are not modified'

$list = @(New-TitleCandidate -Title 'Audio Cd' -Source 'disc label')
$pick = Read-TitleChoice -Candidates $list -OtherDrives @($two[0]) -ReadInput (New-InputQueue @('d', '1')) 6>$null
Assert-Equal 'switch-drive|G:' "$($pick.Source)|$($pick.SwitchDrive)" '[D] in the title list returns the drive to switch to'
$pick = Read-TitleChoice -Candidates $list -OtherDrives @($two[0]) -ReadInput (New-InputQueue @('d', '', '')) 6>$null
Assert-Equal 'disc label' $pick.Source '[D] then Enter goes back to the list'
$pick = Read-TitleChoice -Candidates $list -ReadInput (New-InputQueue @('d', '')) 6>$null
Assert-Equal 'disc label' $pick.Source '[D] with no other drives is just an invalid choice'
$pick = Read-TitleChoice -Candidates @() -OtherDrives @($two[0]) -ReadInput (New-InputQueue @('d', '1')) 6>$null
Assert-Equal 'switch-drive|G:' "$($pick.Source)|$($pick.SwitchDrive)" 'no suggestions: [D] still offered instead of jumping to typing'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
