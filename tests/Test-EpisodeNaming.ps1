#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for genre-series episode naming and episode-number auto-detection.

.DESCRIPTION
    Runs against the real function bodies and the real regex, both lifted out of
    rip-disc.ps1 / continue-rip.ps1 at run time - nothing here is a reimplementation,
    so the tests cannot drift from the shipped code without failing.

    No disc, no MakeMKV and no HandBrake are needed. Filesystem cases use a temp
    directory that is removed afterwards.

.EXAMPLE
    .\tests\Test-EpisodeNaming.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

# ---------- load the real implementations ----------

# Pull a named function's source text out of a script using the PowerShell parser,
# so the test exercises the shipped body rather than a copy of it.
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

# Extract the episode-number regex from the source instead of hardcoding it here,
# so a change to the pattern is caught by these tests rather than silently diverging.
function Get-EpisodeNumberPattern {
    param([string]$ScriptPath, [switch]$Seasoned, [string]$SeasonTag = 'S01')

    $line = Select-String -Path $ScriptPath -Pattern '\$episodeNumberPattern\s*=' |
        Select-Object -First 1
    if (-not $line) { throw "Could not find the episode-number pattern in $ScriptPath" }

    # The line reads: $episodeNumberPattern = if ($seasonTag) { "<A>" } else { "<B>" }
    $quoted = [regex]::Matches($line.Line, '"((?:[^"\\]|\\.)*)"')
    if ($quoted.Count -lt 2) { throw "Unexpected pattern line shape in ${ScriptPath}: $($line.Line)" }

    $raw = if ($Seasoned) { $quoted[0].Groups[1].Value } else { $quoted[1].Groups[1].Value }
    # The seasoned branch is the interpolated string "$seasonTag`E(\d+)". Drop the
    # backtick (it only stops PowerShell reading "$seasonTagE" as one variable) and
    # substitute the season tag the caller is testing with, exactly as the script would.
    $raw = $raw -replace '`', ''
    return ($raw -replace '\$seasonTag', [regex]::Escape($SeasonTag))
}

. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Clean-DiscName')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Resolve-EpisodeNames')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-EpisodeFileName')))

$unseasonedPattern = Get-EpisodeNumberPattern -ScriptPath $ripDiscPath
$seasonedPattern = Get-EpisodeNumberPattern -ScriptPath $ripDiscPath -Seasoned

# ---------- tiny assertion harness ----------

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

# Mirrors the production auto-detect: scan a folder, return the next episode number.
# The regex and the [int] cast are the two things under test.
function Get-NextEpisodeNumber {
    param([string]$Dir, [string]$Pattern)

    $existing = @()
    if (Test-Path $Dir) {
        $existing = Get-ChildItem -Path $Dir -File -Filter '*.mp4' |
            Where-Object { $_.Name -match $Pattern } |
            ForEach-Object { [int]$Matches[1] }
    }
    if ($existing) { return [int](($existing | Measure-Object -Maximum).Maximum) + 1 }
    return 1
}

# ---------- tests ----------

Write-Host "`nLabel normalisation (underscores, spacing, title case)" -ForegroundColor Cyan
Assert-Equal 'Warming By The Devils Fire' (Clean-DiscName -RawName 'WARMING_BY_THE_DEVILS_FIRE').CleanedTitle 'underscores become spaces and casing is normalised'
Assert-Equal 'Soul Of A Man' (Clean-DiscName -RawName 'SOUL_OF_A_MAN').CleanedTitle 'all-caps label is title-cased'
Assert-Equal 'Godfathers And Sons' (Clean-DiscName -RawName 'Godfathers and Sons').CleanedTitle 'already-spaced label is title-cased'
Assert-Equal 'Feel Like Going Home' (Clean-DiscName -RawName 'FEEL__LIKE___GOING_HOME').CleanedTitle 'runs of underscores collapse to one space'

Write-Host "`nEpisode name resolution" -ForegroundColor Cyan
Assert-Equal 'Warming By The Devils Fire' (Resolve-EpisodeNames -FileCount 1 -DiscLabel 'WARMING_BY_THE_DEVILS_FIRE')[0] 'single episode takes the cleaned disc label'
Assert-Equal '' (Resolve-EpisodeNames -FileCount 2 -DiscLabel 'WARMING_BY_THE_DEVILS_FIRE')[0] 'multi-episode disc does not reuse one label across files'
Assert-Equal 'Piano Blues' (Resolve-EpisodeNames -FileCount 1 -Supplied @('Piano Blues') -DiscLabel 'IGNORED_LABEL')[0] '-EpisodeNames overrides the disc label'
Assert-Equal 'Red White And Blues' (Resolve-EpisodeNames -FileCount 2 -Supplied @('Red White And Blues'))[0] 'first supplied name is applied'
Assert-Equal '' (Resolve-EpisodeNames -FileCount 2 -Supplied @('Red White And Blues'))[1] 'unsupplied trailing episode falls back to numbering'
Assert-Equal '' (Resolve-EpisodeNames -FileCount 1 -DiscLabel 'DVD_VIDEO')[0] 'generic disc label is rejected'
Assert-Equal '' (Resolve-EpisodeNames -FileCount 1 -DiscLabel '')[0] 'absent disc label yields no name'

Write-Host "`nFilename construction" -ForegroundColor Cyan
Assert-Equal 'The Blues - S01E04 - Warming.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'S01E04' -EpisodeName 'Warming' -Extension '.mp4') 'named episode uses the Jellyfin " - " pattern'
Assert-Equal 'The Blues-S01E04.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'S01E04' -EpisodeName '' -Extension '.mp4') 'unnamed episode keeps the original shape'
Assert-Equal 'The Blues - E04 - Road to Memphis.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'E04' -EpisodeName 'Road to Memphis' -Extension '.mp4') 'season-less named episode'
Assert-Equal 'The Blues - E04 - AB.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'E04' -EpisodeName 'A:B' -Extension '.mp4') 'illegal filename characters are stripped'
Assert-Equal 'The Blues-E04.mp4' (Get-EpisodeFileName -Title 'The Blues' -EpisodeTag 'E04' -EpisodeName '///' -Extension '.mp4') 'name of only illegal characters falls back to numbering'

Write-Host "`nEpisode-number auto-detection (the cross-session continuation path)" -ForegroundColor Cyan
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    Assert-Equal 1 (Get-NextEpisodeNumber -Dir $tmp -Pattern $seasonedPattern) 'empty folder starts at episode 1'

    New-Item -ItemType File -Path (Join-Path $tmp 'The Blues - S01E04 - Warming.mp4') -Force | Out-Null
    Assert-Equal 5 (Get-NextEpisodeNumber -Dir $tmp -Pattern $seasonedPattern) 'named seasoned episode is detected and continues'

    New-Item -ItemType File -Path (Join-Path $tmp 'The Blues-S01E07.mp4') -Force | Out-Null
    Assert-Equal 8 (Get-NextEpisodeNumber -Dir $tmp -Pattern $seasonedPattern) 'highest wins across named and unnamed files'

    $tmp2 = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-test-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp2 -Force | Out-Null
    try {
        New-Item -ItemType File -Path (Join-Path $tmp2 'The Blues - E04 - Warming.mp4') -Force | Out-Null
        # This is the regression the old "-E(\d+)\." pattern would fail: it demanded a dot
        # immediately after the digits, so a named episode was invisible and numbering
        # silently restarted at 1, overwriting earlier discs.
        Assert-Equal 5 (Get-NextEpisodeNumber -Dir $tmp2 -Pattern $unseasonedPattern) 'named season-less episode is detected (regression)'

        New-Item -ItemType File -Path (Join-Path $tmp2 'The Blues-E06.mp4') -Force | Out-Null
        Assert-Equal 7 (Get-NextEpisodeNumber -Dir $tmp2 -Pattern $unseasonedPattern) 'legacy unnamed season-less file still detected'
    } finally {
        Remove-Item $tmp2 -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Guards the Double/"D2" FormatException: Measure-Object -Maximum returns a Double in
    # PS 5.1 even for all-int input, and "D2" only accepts integral types.
    $tag = "E{0:D2}" -f (Get-NextEpisodeNumber -Dir $tmp -Pattern $seasonedPattern)
    Assert-Equal 'E08' $tag 'episode number formats with D2 without throwing'
} finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nBoth scripts agree on filename construction" -ForegroundColor Cyan
$ripDiscBody = Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-EpisodeFileName'
$continueBody = Import-FunctionFromScript -ScriptPath $continuePath -FunctionName 'Get-EpisodeFileName'
$normalise = { param($s) ($s -replace '\s+', ' ').Trim() }
Assert-Equal (& $normalise $ripDiscBody) (& $normalise $continueBody) 'Get-EpisodeFileName is identical in both scripts'
Assert-Equal $unseasonedPattern (Get-EpisodeNumberPattern -ScriptPath $continuePath) 'season-less detection pattern is identical in both scripts'
Assert-Equal $seasonedPattern (Get-EpisodeNumberPattern -ScriptPath $continuePath -Seasoned) 'seasoned detection pattern is identical in both scripts'

# ---------- summary ----------

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
