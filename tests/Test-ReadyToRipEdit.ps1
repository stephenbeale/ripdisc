<#
.SYNOPSIS
    Tests for [E] Edit at the "Ready to rip" prompt (Read-ReadyToRipEdit in rip-disc.ps1):
    each field asked in turn with Enter keeping it, and the restart parameters it returns.
    Scripted answers via the -ReadInput seam - no disc, MakeMKV or console needed.
#>
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$null)
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Read-ReadyToRipEdit')))

$script:Passed = 0; $script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ("$Expected" -ceq "$Actual") { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because - expected '$Expected', got '$Actual'" -ForegroundColor Red }
}

# Answers in order; the prompts seen are recorded
function New-Answers([string[]]$Answers) {
    $script:Prompts = [System.Collections.ArrayList]::new()
    $script:Queue = [System.Collections.Queue]::new()
    foreach ($a in $Answers) { $script:Queue.Enqueue($a) }
    return { param($p) [void]$script:Prompts.Add($p); if ($script:Queue.Count) { $script:Queue.Dequeue() } else { "" } }
}

# The 2026-11-24 case: Michael Collins Side A was about to rip as "Michael Collins (1996)-extras"
$movie = @{ Title = 'Michael Collins (1996)-extras'; Year = '1996'; Disc = 1; Extras = $false; Season = 0; Series = $false; OutputDrive = 'F:'; Drive = 'G:' }
$bound = @{ OutputDrive = 'F:'; Documentary = $true; DriveIndex = 1; title = 'old' }

Write-Host "`nMovie: every field asked, in order" -ForegroundColor Cyan
$p = Read-ReadyToRipEdit -Current $movie -BoundParams $bound -ReadInput (New-Answers @()) 6>$null
Assert-Equal 'Title [Michael Collins (1996)-extras]|Year [1996] (YYYY, or - for none)|Disc number [1]|Extras only - straight into the extras folder? (y/n) [N]|Output drive [F:]|Disc drive [G:]' ($script:Prompts -join '|') 'title, year, disc, extras, output drive, disc drive'
Assert-Equal $true ($null -eq $p) 'Enter everywhere: nothing changed, no restart'

Write-Host "`nMovie: change title, keep the rest" -ForegroundColor Cyan
$p = Read-ReadyToRipEdit -Current $movie -BoundParams $bound -ReadInput (New-Answers @('Michael Collins Side A', '-')) 6>$null
Assert-Equal 'Michael Collins Side A' $p.title 'new title'
Assert-Equal $false $p.ContainsKey('Year') 'year "-" = none: no -Year'
Assert-Equal 1 $p.Disc 'disc kept'
Assert-Equal $false $p.ContainsKey('Extras') 'not extras'
Assert-Equal 'F:|G:' "$($p.OutputDrive)|$($p.Drive)" 'drives kept, -Drive set'
Assert-Equal $false $p.ContainsKey('DriveIndex') '-DriveIndex dropped (it named the old drive)'
Assert-Equal $true $p.Documentary 'other launch switches (the library) carried over'

Write-Host "`nMovie: extras for Disc 2, other year, other drives" -ForegroundColor Cyan
$p = Read-ReadyToRipEdit -Current $movie -BoundParams $bound -ReadInput (New-Answers @('Michael Collins', 'abcd', '1996', '2', 'y', 'e', 'd:')) 6>$null
Assert-Equal 'Michael Collins|1996|2|True|E:|D:' "$($p.title)|$($p.Year)|$($p.Disc)|$($p.Extras)|$($p.OutputDrive)|$($p.Drive)" 'all fields changed; a bad year is asked again'

Write-Host "`nSeries: season asked, no year or extras question" -ForegroundColor Cyan
$series = @{ Title = 'Silicon Valley'; Year = ''; Disc = 2; Extras = $false; Season = 2; Series = $true; OutputDrive = 'F:'; Drive = 'D:' }
$p = Read-ReadyToRipEdit -Current $series -BoundParams @{ OutputDrive = 'F:' } -ReadInput (New-Answers @('', '3', '1')) 6>$null
Assert-Equal 'Title [Silicon Valley]|Season [2]|Disc number [2]|Output drive [F:]|Disc drive [D:]' ($script:Prompts -join '|') 'series fields'
Assert-Equal 'True|3|1' "$($p.Series)|$($p.Season)|$($p.Disc)" '-Series kept, new season and disc'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
