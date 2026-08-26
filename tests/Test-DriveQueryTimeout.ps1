#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for the drive-query-hang fixes: duplicate-index drive matching and the
    process-timeout-and-kill mechanism.

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 at run time via the
    PowerShell AST parser, so the tests cannot drift from the shipped code without failing.

    Select-MatchedDrive is pure logic and tested against fixture data reproducing the exact
    duplicate-index scenario seen live (a drive that reconnects mid-session can appear twice in
    MakeMKV's own drive list under the same disc:N index).

    Wait-ProcessWithTimeout is inherently a real-process integration point (it polls a live
    System.Diagnostics.Process and kills it on timeout), so it is exercised against real short-
    lived child processes rather than mocked - a genuinely hung process (simulated with a
    Start-Sleep child) and a genuinely fast one, both bounded to short test timeouts so the
    suite stays quick.

.EXAMPLE
    .\tests\Test-DriveQueryTimeout.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

if (-not (Test-Path $ripDiscPath)) { throw "Cannot find $ripDiscPath - run this from inside the repo." }

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

. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Select-MatchedDrive')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Wait-ProcessWithTimeout')))

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
    Assert-Equal $true $Condition $Because
}

Write-Host "`nSelect-MatchedDrive - normal, non-duplicate cases" -ForegroundColor Cyan

$normalLines = @(
    [PSCustomObject]@{ Index = 0; Name = "BD-RE HL-DT-ST BD-RE BU40N 1.05 MO4P6N95940"; DiscName = "DVDVolume"; Letter = "D:"; Busy = $false }
    [PSCustomObject]@{ Index = 1; Name = "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933"; DiscName = "BURN"; Letter = "H:"; Busy = $false }
)
$r1 = Select-MatchedDrive -DrvLines $normalLines -MatchedIndex 1 -DriveLetter "H:"
Assert-Equal 1 (@($r1).Count) "returns exactly one entry when indices are unique"
Assert-Equal "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933" $r1.Name "returns the correct drive's Name"
Assert-Equal "BURN" $r1.DiscName "returns the correct drive's DiscName"

Write-Host "`nSelect-MatchedDrive - the live duplicate-index bug" -ForegroundColor Cyan

# Reproduces exactly what was seen live: a drive that reconnected mid-session shows up twice in
# MakeMKV's own drive list under the SAME index (1) - once keyed by drive letter, once by the
# raw device path MakeMKV fell back to.
$duplicateLines = @(
    [PSCustomObject]@{ Index = 0; Name = "BD-RE HL-DT-ST BD-RE BU40N 1.05 MO4P6N95940"; DiscName = "DVDVolume"; Letter = "D:"; Busy = $false }
    [PSCustomObject]@{ Index = 1; Name = "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933"; DiscName = "BURN"; Letter = "H:"; Busy = $false }
    [PSCustomObject]@{ Index = 1; Name = "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933"; DiscName = "BURN"; Letter = "\\Device\\CdRom3"; Busy = $false }
)

# The old (buggy) behaviour, kept here only to prove the fixture actually reproduces the bug -
# not something the shipped code does any more.
$oldBuggyMatch = $duplicateLines | Where-Object { $_.Index -eq 1 }
Assert-Equal 2 (@($oldBuggyMatch).Count) "fixture check: matching by index alone finds both duplicate entries (this is the bug being fixed)"

$r2 = Select-MatchedDrive -DrvLines $duplicateLines -MatchedIndex 1 -DriveLetter "H:"
Assert-Equal 1 (@($r2).Count) "returns exactly ONE entry even when the index is duplicated"
Assert-Equal "H:" $r2.Letter "picks the entry matching the target drive letter, not the raw-device-path duplicate"
Assert-Equal "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.01 K0MMB391933" $r2.Name "Name is a plain scalar string, not a doubled-up value from a 2-element array"

$r3 = Select-MatchedDrive -DrvLines $duplicateLines -MatchedIndex 1 -DriveLetter "\\Device\\CdRom3"
Assert-Equal "\\Device\\CdRom3" $r3.Letter "the raw-device-path entry is still selectable when IT is the target letter"

Write-Host "`nSelect-MatchedDrive - fallback when no letter match exists" -ForegroundColor Cyan

$r4 = Select-MatchedDrive -DrvLines $normalLines -MatchedIndex 1 -DriveLetter "Z:"
Assert-Equal "H:" $r4.Letter "falls back to the first same-index entry when no entry matches the given letter (defensive path)"

$r5 = Select-MatchedDrive -DrvLines $normalLines -MatchedIndex 99 -DriveLetter "H:"
Assert-True ($null -eq $r5) "returns `$null when no entry matches the index at all"

Write-Host "`nWait-ProcessWithTimeout - a fast process exits on its own" -ForegroundColor Cyan

$fastProc = New-Object System.Diagnostics.Process
$fastProc.StartInfo.FileName = "powershell.exe"
$fastProc.StartInfo.Arguments = "-NoProfile -Command `"exit 0`""
$fastProc.StartInfo.UseShellExecute = $false
$fastProc.StartInfo.CreateNoWindow = $true
$fastProc.Start() | Out-Null
$fastSw = [System.Diagnostics.Stopwatch]::StartNew()
$fastResult = Wait-ProcessWithTimeout -Process $fastProc -TimeoutSec 15
Assert-True $fastResult "returns `$true when the process exits on its own within the timeout"
Assert-True ($fastSw.Elapsed.TotalSeconds -lt 10) "does not wait for the full timeout when the process already exited"
Assert-True $fastProc.HasExited "the process has genuinely exited"

Write-Host "`nWait-ProcessWithTimeout - a hung process is killed on timeout" -ForegroundColor Cyan

$hungProc = New-Object System.Diagnostics.Process
$hungProc.StartInfo.FileName = "powershell.exe"
$hungProc.StartInfo.Arguments = "-NoProfile -Command `"Start-Sleep -Seconds 120`""  # simulates a hang
$hungProc.StartInfo.UseShellExecute = $false
$hungProc.StartInfo.CreateNoWindow = $true
$hungProc.Start() | Out-Null
$hungSw = [System.Diagnostics.Stopwatch]::StartNew()
$hungResult = Wait-ProcessWithTimeout -Process $hungProc -TimeoutSec 2
$hungElapsed = $hungSw.Elapsed.TotalSeconds
Assert-Equal $false $hungResult "returns `$false when the process has to be killed"
Assert-True ($hungElapsed -lt 10) "bounds the wait to roughly the timeout, not the process's full 120s runtime"
Start-Sleep -Milliseconds 500  # give the OS a moment to finish tearing the killed process down
Assert-True $hungProc.HasExited "the hung process was actually killed, not just abandoned"

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
