#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for movie / genre-film folder naming ("Slow Dance" or "Slow Dance (2014)"):
    the shared helpers in Load-Config.ps1, the year prompt in rip-disc.ps1, the -Year
    hand-off to continue-rip.ps1, and that both scripts build movie folders from
    $script:MediaFolderName.

.DESCRIPTION
    Lifts the real functions out of the scripts with the AST parser, so nothing here is
    a reimplementation. Filesystem cases use a temp directory removed afterwards. No
    network, disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-MovieFolderNaming.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$configPath = Join-Path $repoRoot 'Load-Config.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'
$setupPath = Join-Path $repoRoot 'setup.ps1'

foreach ($p in @($configPath, $ripDiscPath, $continuePath, $setupPath)) {
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
    $fn = (Get-ScriptAst $ScriptPath).FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

function Get-ScriptParameterNames {
    param([string]$ScriptPath)
    return @((Get-ScriptAst $ScriptPath).ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}

foreach ($name in @('Get-MovieParentFolder', 'Get-MediaFolderName', 'Find-ExistingMediaFolder')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
}
foreach ($name in @('New-TitleCandidate', 'Get-TitleCandidateUrl', 'Format-TitleCandidate', 'Show-TitleCandidates',
                    'Read-FolderYear', 'Get-ContinueRipCommand')) {
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

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-folder-naming-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    # ---------------------------------------------------------------------------
    Write-Host "`nFolder name and parent folder" -ForegroundColor Cyan

    Assert-Equal 'Slow Dance (2014)' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '2014') 'year appended in brackets'
    Assert-Equal 'Slow Dance' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '') 'no year, no change'
    Assert-Equal 'Slow Dance' (Get-MediaFolderName -SafeTitle 'Slow Dance' -Year '14') 'a non-4-digit year is ignored'
    Assert-Equal 'Surf' (Get-MovieParentFolder -Surf) '-Surf films live under Surf'
    Assert-Equal 'Documentaries' (Get-MovieParentFolder -Documentary -Bluray) 'a genre wins over -Bluray (same order as the scripts)'
    Assert-Equal 'Bluray' (Get-MovieParentFolder -Bluray) 'Blu-ray movies under Bluray'
    Assert-Equal 'DVDs' (Get-MovieParentFolder) 'plain movies under DVDs'
    Assert-True ('Slow Dance (2014)-title_t00.mp4' -like ('Slow Dance (2014)' + '-*')) 'parentheses are not wildcards: the movie-mode prefix check still matches'

    # ---------------------------------------------------------------------------
    Write-Host "`nReusing an existing folder" -ForegroundColor Cyan

    $surf = Join-Path $tempRoot 'Surf'
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance')) 'parent folder missing: nothing found'
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance (2014)') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance 2') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dancers (2010)') -Force | Out-Null
    Assert-Equal 'Slow Dance (2014)' (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance') 'Disc 2 finds the "(2014)" folder Disc 1 made; similar names ignored'
    Assert-Equal 'Slow Dance (2014)' (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'slow dance') 'match ignores case, like Windows'
    New-Item -ItemType Directory -Path (Join-Path $surf 'Slow Dance') -Force | Out-Null
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $surf -SafeTitle 'Slow Dance')) 'two candidates (plain and dated): no guess'
    $dvds = Join-Path $tempRoot 'DVDs'
    New-Item -ItemType Directory -Path (Join-Path $dvds 'W_ (2008)') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dvds 'Heat (Director''s Cut)') -Force | Out-Null
    Assert-Equal 'W_ (2008)' (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'W_') 'sanitized title matched literally'
    Assert-Equal 'Heat (Director''s Cut)' (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'Heat (Director''s Cut)') 'regex characters in the title are escaped'
    Assert-True ($null -eq (Find-ExistingMediaFolder -ParentDir $dvds -SafeTitle 'Heat')) '"Heat (Director''s Cut)" is not taken as a year folder for "Heat"'

    # ---------------------------------------------------------------------------
    Write-Host "`nYear prompt" -ForegroundColor Cyan

    $hits = @(
        (New-TitleCandidate -Title 'Slow Dance' -Year '2014' -MediaType 'movie' -Kind 'Short' -ImdbId 'tt3453042' -Source 'IMDb'),
        (New-TitleCandidate -Title 'Slow Dance' -Source 'disc label'),
        (New-TitleCandidate -Title 'Slow Dance' -Year '2005' -MediaType 'tv' -Kind 'TV mini-series' -ImdbId 'tt0481050' -Source 'IMDb'))
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('')) 6>$null)) 'Enter: no year'
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('s')) 6>$null)) 'S: no year'
    Assert-True ($null -eq (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @()) 6>$null)) 'end of input: no year (the rip is not aborted)'
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('1')) 6>$null) 'row 1 gives its year'
    Assert-Equal '2005' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('2')) 6>$null) 'rows without a year are left out, so row 2 is the 2005 hit'
    Assert-Equal '2013' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('2013')) 6>$null) 'a typed year is used as-is'
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates $hits -ReadInput (New-InputQueue @('9', 'x', '1')) 6>$null) 'invalid answers re-ask'
    Assert-Equal '1999' (Read-FolderYear -Title 'Slow Dance' -Candidates @() -ReadInput (New-InputQueue @('1999')) 6>$null) 'no hits: a year can still be typed'
    $script:Searched = @()
    $searchStub = { param($q) $script:Searched += $q; @(New-TitleCandidate -Title 'Slow Dance' -Year '2014' -Kind 'Short' -Source 'IMDb') }
    Assert-Equal '2014' (Read-FolderYear -Title 'Slow Dance' -Candidates @() -Search $searchStub -ReadInput (New-InputQueue @('r', 'slow dance surf', '1')) 6>$null) 'R searches again and the new rows can be picked'
    Assert-Equal 'slow dance surf' ($script:Searched -join ';') 'the typed search text is used'

    # ---------------------------------------------------------------------------
    Write-Host "`nHand-off to continue-rip.ps1" -ForegroundColor Cyan

    Assert-True ((Get-ScriptParameterNames $ripDiscPath) -contains 'Year') 'rip-disc.ps1 has -Year'
    Assert-True ((Get-ScriptParameterNames $continuePath) -contains 'Year') 'continue-rip.ps1 has -Year'
    $cmd = Get-ContinueRipCommand -Title 'Slow Dance' -RemainingStepNumbers @(2) -Surf -Disc 1 -StartEpisode 1 -EpisodeNames @() -Year '2014'
    Assert-Equal '.\continue-rip.ps1 -title "Slow Dance" -FromStep handbrake -Surf -Year 2014' $cmd 'retry command carries the folder year'
    $cmd = Get-ContinueRipCommand -Title 'Slow Dance' -RemainingStepNumbers @(2) -Surf -Disc 1 -StartEpisode 1 -EpisodeNames @()
    Assert-Equal '.\continue-rip.ps1 -title "Slow Dance" -FromStep handbrake -Surf' $cmd 'no year, no -Year'
    $cmd = Get-ContinueRipCommand -Title 'Fargo' -RemainingStepNumbers @(3) -Series -Season 1 -Disc 1 -StartEpisode 1 -EpisodeNames @() -Year '2014'
    Assert-True ($cmd -notmatch '-Year') 'series commands never carry -Year (series folders are unchanged)'

    # ---------------------------------------------------------------------------
    Write-Host "`nBoth scripts build movie folders from the chosen name" -ForegroundColor Cyan

    foreach ($path in @($ripDiscPath, $continuePath)) {
        $text = Get-Content -LiteralPath $path -Raw
        $leaf = Split-Path $path -Leaf
        foreach ($folder in @('Documentaries', 'Tutorials', 'Fitness', 'Music', 'Surf', 'Bluray', 'DVDs')) {
            Assert-True ($text.Contains("`$finalOutputDir = `"`$outputDriveLetter\$folder\`$(`$script:MediaFolderName)`"")) "$leaf`: $folder folder uses the chosen name"
        }
        Assert-True ($text.Contains('$seriesBaseDir = "$outputDriveLetter\Series\$safeTitle"')) "$leaf`: series folder still the plain title"
    }

    # ---------------------------------------------------------------------------
    Write-Host "`nNever rip a film into another film's folder (Scarface 1932, 2026-10-28)" -ForegroundColor Cyan

    foreach ($name in @('Get-ExistingFeatureFile', 'Get-FeatureNamePattern', 'Get-ImageFileNames')) {
        . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
    }
    $script:AnyFeaturePattern = '-Feature(-BluRay)?\.[^.]+$'
    Assert-True ((Get-Content -LiteralPath $configPath -Raw).Contains("`$script:AnyFeaturePattern = '-Feature(-BluRay)?\.[^.]+$'")) 'Load-Config.ps1 defines the same any-version pattern used here'
    . ([scriptblock]::Create((Import-FunctionFromScript $ripDiscPath 'Get-UniqueFilePath')))
    $dvds = Join-Path $tempRoot 'clash\DVDs'
    New-Item -ItemType Directory -Path (Join-Path $dvds 'Scarface\extras') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $dvds 'Scarface\Scarface-Feature.mp4') -Value 'x'
    Set-Content -LiteralPath (Join-Path $dvds 'Scarface\folder.jpg') -Value 'x'
    $dvdPat = Get-FeatureNamePattern
    Assert-Equal 'Scarface-Feature.mp4' (Get-ExistingFeatureFile -FolderPath (Join-Path $dvds 'Scarface') -Pattern $dvdPat) 'finds an existing Feature'
    Assert-Equal $null (Get-ExistingFeatureFile -FolderPath (Join-Path $dvds 'Scarface (1932)') -Pattern $dvdPat) 'missing folder: no Feature'
    Assert-Equal $null (Get-ExistingFeatureFile -FolderPath (Join-Path $dvds 'Scarface') -Pattern (Get-FeatureNamePattern -Part 1)) 'a Part pattern ignores a Feature'
    Assert-Equal $null (Get-ExistingFeatureFile -FolderPath (Join-Path $dvds 'Scarface') -Pattern (Get-FeatureNamePattern -Bluray)) 'a Blu-ray run does not see the DVD Feature as its own'
    Assert-Equal '-Feature\.[^.]+$' $dvdPat 'DVD main-file pattern'
    Assert-Equal '-Feature-BluRay\.[^.]+$' (Get-FeatureNamePattern -Bluray) 'Blu-ray main-file pattern'
    Assert-Equal '-Part 2\.[^.]+$' (Get-FeatureNamePattern -Part 2 -Bluray) 'a documentary part is named by part, Blu-ray or not'

    # The real folder-choice block from rip-disc.ps1, run with prompts/lookups stubbed and
    # "exit 1" turned into a throw so a refusal can be seen
    $ripText = Get-Content -LiteralPath $ripDiscPath -Raw
    $start = $ripText.IndexOf('# ========== MOVIE FOLDER NAME')
    $end = $ripText.IndexOf('# ========== SERIES EPISODE NUMBERING')
    $folderBlock = [scriptblock]::Create($ripText.Substring($start, $end - $start).Replace('exit 1', "throw 'REFUSED'"))
    function Get-NormalizedDriveLetter { param($d) $d }
    function Get-DiscTitleCandidates { @() }
    function Get-TMDbTitleCandidates { @() }
    function Get-ImdbTitleCandidates { @() }
    $outputRoot = Join-Path $tempRoot 'clash'
    function Invoke-FolderChoice {
        param([string]$T, $Pick = $null, [string]$Y = '', [int]$D = 1, [switch]$X, [string]$Naming = 'title', [string]$YearAnswer = '',
              [switch]$BD, [switch]$Doc, [switch]$Part)
        $script:Config_MovieFolderNaming = $Naming
        $script:DocumentaryPart = [bool]$Part
        $script:YearAnswer = $YearAnswer
        function Read-FolderYear { param($Title, $Candidates, $Search) $script:YearAnswer }
        $title = $T; $safeTitle = $T; $picked = $Pick; $Year = $Y; $Disc = $D; $Extras = [bool]$X
        $Series = $false; $OutputDrive = $outputRoot
        $Tutorial = $Fitness = $Music = $Surf = $false
        $Documentary = [bool]$Doc; $Bluray = [bool]$BD
        try { . $folderBlock 6>$null; return $script:MediaFolderName } catch { return "$($_.Exception.Message)" }
    }
    $pick1932 = [pscustomobject]@{ Title = 'Scarface'; Year = 1932 }
    Assert-Equal 'Scarface (1932)' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932) 'THE BUG: picked Scarface (1932) goes to "Scarface (1932)", not the 1983 film''s folder'
    Assert-Equal 'Scarface (1932)' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932 -Naming 'title-year') 'same with title-year naming'
    Assert-Equal 'Scarface (1983)' (Invoke-FolderChoice -T 'Scarface' -YearAnswer '1983') 'no picked year: asks for one and uses it'
    Assert-Equal 'REFUSED' (Invoke-FolderChoice -T 'Scarface' -YearAnswer '') 'no year given: refuses before ripping'
    Assert-Equal 'Scarface (1983)' (Invoke-FolderChoice -T 'Scarface' -Y '1983') '-Year to a free folder is used'
    New-Item -ItemType Directory -Path (Join-Path $dvds 'Scarface (1932)') -Force | Out-Null
    Assert-Equal 'Scarface (1932)' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932) 'an empty "Scarface (1932)" folder is reused'
    Set-Content -LiteralPath (Join-Path $dvds 'Scarface (1932)\Scarface (1932)-Feature.mp4') -Value 'x'
    Assert-Equal 'Scarface (1932)' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932 -D 2) 'Disc 2 of the 1932 film goes to its own folder'
    Assert-Equal 'Scarface (1932)' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932 -X) '-Extras of the 1932 film goes to its own folder'
    Assert-Equal 'REFUSED' (Invoke-FolderChoice -T 'Scarface' -Pick $pick1932) 'both folders hold a Feature: refuses before ripping'
    Assert-Equal 'REFUSED' (Invoke-FolderChoice -T 'Scarface' -Y '1932') '-Year to a folder that holds a Feature: refuses'
    Assert-Equal 'Slow Dance' (Invoke-FolderChoice -T 'Slow Dance') 'no existing folder: unchanged plain name'

    Write-Host "`nBlu-ray and documentary names are left alone" -ForegroundColor Cyan
    $docs = Join-Path $outputRoot 'Documentaries'
    New-Item -ItemType Directory -Path (Join-Path $docs 'Cream') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $docs 'Cream\Cream-Feature.mp4') -Value 'x'
    Assert-Equal 'Cream' (Invoke-FolderChoice -T 'Cream' -Doc -BD) 'Blu-ray version goes beside the DVD Feature (two Jellyfin versions)'
    Assert-Equal 'REFUSED' (Invoke-FolderChoice -T 'Cream' -Doc) 'a second DVD Feature is still refused'
    New-Item -ItemType Directory -Path (Join-Path $docs 'Metal') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $docs 'Metal\Metal-Part 1.mp4') -Value 'x'
    Assert-Equal 'Metal' (Invoke-FolderChoice -T 'Metal' -Doc -Part -D 2) 'documentary Part 2 goes beside Part 1'
    Assert-Equal 'REFUSED' (Invoke-FolderChoice -T 'Metal' -Doc -Part -D 1) 'a second Part 1 is refused'
    $bluray = Join-Path $outputRoot 'Bluray'
    New-Item -ItemType Directory -Path (Join-Path $bluray 'Scarface') -Force | Out-Null
    Assert-Equal 'Scarface' (Invoke-FolderChoice -T 'Scarface' -BD) 'Blu-ray movies use the Bluray folder; the DVDs Scarface does not clash'

    # The real Step 3 main-file patterns from both scripts
    foreach ($path in @($ripDiscPath, $continuePath)) {
        $leaf = Split-Path $path -Leaf
        $text = Get-Content -LiteralPath $path -Raw
        $s = $text.IndexOf('# Naming of the main file')
        $e = $text.IndexOf('$extrasDir = Join-Path $finalOutputDir "extras"', $s)
        $patBlock = [scriptblock]::Create($text.Substring($s, $e - $s))
        foreach ($case in @(@{ BD = $false; Part = $false; D = 1; Own = 'Cream-Feature.mp4'; Other = 'Cream-Feature-BluRay.mp4' },
                            @{ BD = $true; Part = $false; D = 1; Own = 'Cream-Feature-BluRay.mp4'; Other = 'Cream-Feature.mp4' },
                            @{ BD = $false; Part = $true; D = 2; Own = 'Metal-Part 2.mp4'; Other = 'Metal-Part 1.mp4' })) {
            $Bluray = $case.BD; $Disc = $case.D; $script:DocumentaryPart = $case.Part
            . $patBlock
            $label = "$leaf $(if ($case.Part) { 'part' } elseif ($case.BD) { 'Blu-ray' } else { 'DVD' })"
            Assert-True ($case.Own -match $featurePattern) "$label`: its own main file counts as done"
            Assert-True ($case.Other -notmatch $featurePattern) "$label`: $($case.Other) is not taken for its own"
            Assert-True ($case.Other -match $featureExcludePattern -and $case.Own -match $featureExcludePattern) "$label`: neither main file is ever the 'largest' pick or moved to extras"
            Assert-True ('Cream-title_t00.mp4' -notmatch $featureExcludePattern) "$label`: ordinary encodes are still sorted"
        }
    }

    Write-Host "`nExisting artwork is never deleted" -ForegroundColor Cyan
    Assert-Equal 'folder.jpg' ((Get-ImageFileNames -Dir (Join-Path $dvds 'Scarface')) -join ',') 'artwork snapshot lists images only'
    Assert-Equal 0 @(Get-ImageFileNames -Dir (Join-Path $dvds 'Nope')).Count 'missing folder: empty snapshot'
    foreach ($path in @($ripDiscPath, $continuePath)) {
        $text = Get-Content -LiteralPath $path -Raw
        $leaf = Split-Path $path -Leaf
        Assert-True ($text.Contains('$script:PreexistingImages = Get-ImageFileNames -Dir $finalOutputDir')) "$leaf`: snapshots artwork before the run"
        Assert-True ($text.Contains('-and $script:PreexistingImages -notcontains $_.Name }')) "$leaf`: Step 3 deletes only images that are new"
        Assert-True (([regex]::Matches($text, "Extension -match '\\.\(jpg\|jpeg\|png\|gif\|bmp\)\$' \}")).Count -eq 0) "$leaf`: no unfiltered image delete left"
    }

    Write-Host "`ncontinue-rip.ps1 never organizes into another film's folder" -ForegroundColor Cyan
    $contText = Get-Content -LiteralPath $continuePath -Raw
    $s = $contText.IndexOf("# ========== NEVER ORGANIZE A FILM INTO ANOTHER FILM'S FOLDER")
    $e = $contText.IndexOf('# ========== SERIES EPISODE NUMBERING', $s)
    $contBlock = [scriptblock]::Create($contText.Substring($s, $e - $s).Replace('exit 1', "throw 'REFUSED'"))
    function Enable-ConsoleClose { }
    function Write-Log { param($m) }
    function Invoke-ContinueCheck {
        param([string]$Dir, [int]$Step, [string]$Pattern, [string]$Answer = $null, [bool]$Main = $true)
        $script:ReadHostAnswer = $Answer
        function Read-Host { param($p) $script:ReadHostAnswer }
        $finalOutputDir = $Dir; $StartFromStepNumber = $Step; $featurePattern = $Pattern; $isMainFeatureDisc = $Main
        $script:FeatureIsOurs = $false
        try { . $contBlock 6>$null; return "OK ours=$($script:FeatureIsOurs) images=$($script:PreexistingImages -join ',')" } catch { return "$($_.Exception.Message)" }
    }
    $scar = Join-Path $dvds 'Scarface'
    Assert-Equal 'REFUSED' (Invoke-ContinueCheck -Dir $scar -Step 2 -Pattern $dvdPat) 'from Step 2: an existing Feature is another film - stop'
    Assert-Equal 'REFUSED' (Invoke-ContinueCheck -Dir $scar -Step 3 -Pattern $dvdPat -Answer '') 'from Step 3, Enter: stop'
    Assert-Equal 'REFUSED' (Invoke-ContinueCheck -Dir $scar -Step 3 -Pattern $dvdPat -Answer $null) 'from Step 3, no input: stop'
    Assert-Equal 'OK ours=True images=folder.jpg' (Invoke-ContinueCheck -Dir $scar -Step 3 -Pattern $dvdPat -Answer 'y') 'from Step 3, "y" (interrupted organize): carries on, artwork kept'
    Assert-Equal 'OK ours=False images=folder.jpg' (Invoke-ContinueCheck -Dir $scar -Step 2 -Pattern (Get-FeatureNamePattern -Bluray)) 'Blu-ray continue beside a DVD Feature: no stop'
    Assert-Equal 'OK ours=False images=folder.jpg' (Invoke-ContinueCheck -Dir $scar -Step 4 -Pattern $dvdPat) 'Step 4 only opens the folder: no check'
    Assert-Equal 'OK ours=False images=folder.jpg' (Invoke-ContinueCheck -Dir $scar -Step 2 -Pattern $dvdPat -Main $false) 'Disc 2 / extras into the film folder: no stop'
    Assert-True ($contText.Contains('if ($script:ForeignFeature) {')) 'continue-rip.ps1 Step 3 never sweeps into extras beside a Feature it did not confirm'

    Write-Host "`nHandBrake output never overwrites" -ForegroundColor Cyan
    $enc = Join-Path $tempRoot 'enc'
    New-Item -ItemType Directory -Path $enc | Out-Null
    Set-Content -LiteralPath (Join-Path $enc 'title_t00.mp4') -Value 'old'
    Assert-Equal (Join-Path $enc 'title_t00-1.mp4') (Get-UniqueFilePath -DestDir $enc -FileName 'title_t00.mp4') 'existing name gets -1'
    Assert-Equal (Join-Path $enc 'title_t01.mp4') (Get-UniqueFilePath -DestDir $enc -FileName 'title_t01.mp4') 'free name kept'
    Assert-True ($ripText.Contains('$outputFile = $encodeTargets[$mkv.FullName]')) 'rip-disc.ps1 encodes to the de-clashed paths'
    Assert-True (-not $ripText.Contains('$outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")')) 'rip-disc.ps1 has no raw BaseName.mp4 output left'
    Assert-True ($ripText.Contains('if ($script:ForeignFeature) {')) 'Step 3 never sweeps a new film into extras beside an existing Feature'

    # ---------------------------------------------------------------------------
    Write-Host "`nsetup.ps1 and config" -ForegroundColor Cyan

    $setupText = Get-Content -LiteralPath $setupPath -Raw
    Assert-True ($setupText -match 'movieFolderNaming = \$movieFolderNaming') 'setup.ps1 saves movieFolderNaming'
    Assert-True ($setupText -match '\[6/6\] Movie folder naming' -and $setupText -notmatch '/5\]') 'setup steps renumbered to 6'
    $sample = Get-Content -LiteralPath (Join-Path $repoRoot 'ripdisc-config.sample.json') -Raw | ConvertFrom-Json
    Assert-Equal 'title' $sample.movieFolderNaming 'sample config defaults to plain titles'
    $configText = Get-Content -LiteralPath $configPath -Raw
    Assert-True ($configText -match "-in @\('title', 'title-year'\)") 'Load-Config.ps1 accepts only the two styles (anything else falls back to title)'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
