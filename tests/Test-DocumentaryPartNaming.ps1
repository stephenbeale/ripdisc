#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for "-Documentary -Disc N" part naming ("<Title>-Part N" in the title folder).

.DESCRIPTION
    Lifts the real decision block ($script:DocumentaryPart, $isMainFeatureDisc, $featurePattern,
    $featureSuffix, $featureExcludePattern) and the real Step 3 selection lines out of
    rip-disc.ps1 / continue-rip.ps1 at run time and executes them, so the tests cannot drift
    from the shipped code. No disc, MakeMKV or HandBrake needed; filesystem cases use a temp
    directory that is removed afterwards.

.EXAMPLE
    .\tests\Test-DocumentaryPartNaming.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$scripts = @{ 'rip-disc.ps1' = (Join-Path $repoRoot 'rip-disc.ps1'); 'continue-rip.ps1' = (Join-Path $repoRoot 'continue-rip.ps1') }

$script:Passed = 0; $script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ("$Expected" -ceq "$Actual") { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because" -ForegroundColor Red; Write-Host "        expected: [$Expected]`n        actual:   [$Actual]" -ForegroundColor Red }
}

function Get-Block {
    param([string]$Path, [string]$StartPattern, [string]$EndPattern)
    $lines = Get-Content -LiteralPath $Path
    $s = ($lines | Select-String -Pattern $StartPattern | Select-Object -First 1).LineNumber
    if (-not $s) { throw "Start pattern not found in ${Path}: $StartPattern" }
    $e = $s
    for ($i = $s; $i -le $lines.Count; $i++) { if ($lines[$i - 1] -match $EndPattern) { $e = $i; break } }
    return ($lines[($s - 1)..($e - 1)] -join "`n")
}

# Shared helpers the decision block calls (Load-Config.ps1, dot-sourced by both scripts)
$cfgAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'Load-Config.ps1'), [ref]$null, [ref]$null)
$cfgAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-FeatureNamePattern' }, $true) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
$script:AnyFeaturePattern = '-Feature(-BluRay)?\.[^.]+$'

function Get-Line {
    param([string]$Path, [string]$Pattern)
    $m = Select-String -LiteralPath $Path -Pattern $Pattern -SimpleMatch | Select-Object -First 1
    if (-not $m) { throw "Line not found in ${Path}: $Pattern" }
    return $m.Line.Trim()
}

# Evaluate the decision block with given switch values. PSBoundParameters is faked as a hashtable.
function Invoke-Decision {
    param([string]$Path, [bool]$Documentary = $false, [bool]$Series = $false, [bool]$Extras = $false,
          [bool]$Bluray = $false, [int]$Disc = 1, [bool]$DiscBound = $false)
    $BoundFake = @{}
    if ($DiscBound) { $BoundFake['Disc'] = $Disc }
    $partLine = (Get-Line $Path '$script:DocumentaryPart = ').Replace('$PSBoundParameters', '$BoundFake')
    $block = Get-Block $Path '^\$isMainFeatureDisc = ' '^\$featureExcludePattern = '
    $script:DocumentaryPart = $null
    Invoke-Expression $partLine
    Invoke-Expression $block
    return [pscustomobject]@{ Part = $script:DocumentaryPart; Main = $isMainFeatureDisc; Pattern = $featurePattern; Suffix = $featureSuffix; Exclude = $featureExcludePattern }
}

foreach ($name in $scripts.Keys) {
    $p = $scripts[$name]
    Write-Host "`n[$name] decision block" -ForegroundColor Cyan

    $r = Invoke-Decision $p -Documentary $true -Disc 2 -DiscBound $true
    Assert-Equal $true $r.Part '-Documentary -Disc 2 is a documentary part'
    Assert-Equal $true $r.Main 'a documentary part is organised as a main-file disc (stays in title folder, not extras)'
    Assert-Equal '-Part 2' $r.Suffix 'suffix is -Part <Disc>'

    $r = Invoke-Decision $p -Documentary $true -Disc 1 -DiscBound $true
    Assert-Equal '-Part 1' $r.Suffix 'explicit -Disc 1 gives -Part 1'

    $r = Invoke-Decision $p -Documentary $true
    Assert-Equal $false $r.Part '-Documentary without explicit -Disc keeps the old behaviour'
    Assert-Equal '-Feature' $r.Suffix '... and the -Feature suffix'

    $r = Invoke-Decision $p -Documentary $true -Disc 2 -DiscBound $false
    Assert-Equal $false $r.Main 'Disc>1 not explicitly bound (e.g. auto-detected) still treated as special features, as before'

    $r = Invoke-Decision $p -Documentary $false -Disc 2 -DiscBound $true
    Assert-Equal $false $r.Part 'plain movie with -Disc 2 is NOT a documentary part'
    Assert-Equal $false $r.Main 'plain movie Disc 2 remains a special-features disc'

    $r = Invoke-Decision $p -Documentary $true -Series $true -Disc 2 -DiscBound $true
    Assert-Equal $false $r.Part '-Documentary -Series (genre series) is unaffected'

    $r = Invoke-Decision $p -Documentary $true -Extras $true -Disc 2 -DiscBound $true
    Assert-Equal $false $r.Part '-Documentary -Extras is unaffected'

    $r = Invoke-Decision $p -Bluray $true
    Assert-Equal '-Feature-BluRay' $r.Suffix '-Bluray movie still gets -Feature-BluRay'
}

# Filesystem behaviour using the real selection lines from rip-disc.ps1
Write-Host "`nStep 3 file selection (real lines from rip-disc.ps1)" -ForegroundColor Cyan
$p = $scripts['rip-disc.ps1']
$largestLine = Get-Line $p '$largestFile = Get-ChildItem -File'
$existsLine  = Get-Line $p '$featureExists = Get-ChildItem -File'
$nonFeatLine = Get-Line $p '$nonFeatureVideos = Get-ChildItem -File'

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("docpart-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    Push-Location $tmp
    # Disc 1 already organised; Disc 2 freshly encoded and prefixed with the folder name
    $mk = { param($n, $bytes) $fs = [IO.File]::Create((Join-Path $tmp $n)); $fs.SetLength($bytes); $fs.Close() }
    & $mk 'Metal-Part 1.mp4' 9000
    & $mk 'Metal-t00.mp4' 5000
    & $mk 'Metal-t01.mp4' 100

    $BoundFake = @{ Disc = 2 }
    $Documentary = $true; $Series = $false; $Extras = $false; $Bluray = $false; $Disc = 2
    Invoke-Expression (Get-Line $p '$script:DocumentaryPart = ').Replace('$PSBoundParameters', '$BoundFake')
    Invoke-Expression (Get-Block $p '^\$isMainFeatureDisc = ' '^\$featureExcludePattern = ')

    Invoke-Expression $existsLine
    Assert-Equal $null $featureExists 'Part 1 on disk is not mistaken for Disc 2 already having its Part file'
    Invoke-Expression $largestLine
    Assert-Equal 'Metal-t00.mp4' $largestFile.Name "Disc 1's larger Part 1 file is not picked as Disc 2's largest"
    Assert-Equal 'Metal-Part 2.mp4' ('Metal' + $featureSuffix + $largestFile.Extension) 'new name is <Title>-Part 2.mp4'

    Rename-Item -LiteralPath $largestFile.FullName -NewName 'Metal-Part 2.mp4'
    Invoke-Expression $existsLine
    Assert-Equal 'Metal-Part 2.mp4' $featureExists.Name 'a re-run finds the existing Part 2 and skips renaming'
    Invoke-Expression $nonFeatLine
    Assert-Equal 'Metal-t01.mp4' (@($nonFeatureVideos).Name -join ',') 'only the leftover title goes to extras; Part 1 and Part 2 stay put'
} finally {
    Pop-Location
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# continue-rip command carries -Disc for a documentary part, even Disc 1
Write-Host "`nContinue command" -ForegroundColor Cyan
$src = Get-Content -Raw -LiteralPath $scripts['rip-disc.ps1']
Assert-Equal $true ($src.Contains('if ($Disc -ne 1 -or $script:DocumentaryPart)')) 'rip-disc.ps1 emits -Disc for explicit documentary parts so the resume keeps the naming'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
