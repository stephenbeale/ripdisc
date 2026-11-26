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

Write-Host "`nWait-ProcessWithTimeout - countdown ticks" -ForegroundColor Cyan

$tickProc = New-Object System.Diagnostics.Process
$tickProc.StartInfo.FileName = "powershell.exe"
$tickProc.StartInfo.Arguments = "-NoProfile -Command `"Start-Sleep -Seconds 120`""
$tickProc.StartInfo.UseShellExecute = $false
$tickProc.StartInfo.CreateNoWindow = $true
$tickProc.Start() | Out-Null
$ticks = [System.Collections.Generic.List[int]]::new()
$tickResult = Wait-ProcessWithTimeout -Process $tickProc -TimeoutSec 3 -OnTick { param($s) $ticks.Add($s) }
Assert-Equal $false $tickResult "still times out with -OnTick"
Assert-Equal '1,2' (($ticks | Select-Object -First 2) -join ',') "-OnTick gets each whole second once, in order"
Assert-True ($ticks.Count -le 3) "no duplicate ticks for the same second"
Start-Sleep -Milliseconds 500

Write-Host "`nDrive-list cache - only a real makemkvcon's clean output, keyed by its path" -ForegroundColor Cyan
foreach ($fn in 'Get-DriveCacheHeader', 'Test-CacheableMakeMkvPath', 'Read-DriveCache', 'Write-DriveCache') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $fn)))
}
$cacheDir = Join-Path ([IO.Path]::GetTempPath()) "ripdisc-cache-test-$PID"
New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
$cacheFile = Join-Path $cacheDir 'makemkv-drive-cache.txt'
$realMkv = 'C:\Program Files (x86)\MakeMKV\makemkvcon64.exe'
$good = @('MSG:1005,0,1,"started"', 'DRV:0,2,999,0,"GP75N","DISC","D:"')
try {
    # The 2026-11-22 incident: a fake makemkvcon's drive list must never be cached
    Assert-Equal 'not a real makemkvcon (C:\tmp\fake-makemkvcon.cmd)' (Write-DriveCache -CacheFile $cacheFile -MakeMkvPath 'C:\tmp\fake-makemkvcon.cmd' -Lines $good) "a fake makemkvcon's output is not cached"
    Assert-True (-not (Test-Path $cacheFile)) "no cache file written for the fake"
    Assert-Equal 'read errors' (Write-DriveCache -CacheFile $cacheFile -MakeMkvPath $realMkv -Lines ($good + 'MSG:2003,0,3,"Error"')) "read errors (MSG:2003) are not cached"
    Assert-Equal '' (Write-DriveCache -CacheFile $cacheFile -MakeMkvPath $realMkv -Lines $good) "a real makemkvcon64.exe's clean output is cached"
    $back = Read-DriveCache -CacheFile $cacheFile -MakeMkvPath $realMkv
    Assert-Equal ($good -join '|') (@($back) -join '|') "read back without the header line"
    Assert-True ($null -eq (Read-DriveCache -CacheFile $cacheFile -MakeMkvPath 'C:\Other\makemkvcon64.exe')) "a cache written by a different makemkvcon path is ignored"
    Assert-True ($null -eq (Read-DriveCache -CacheFile $cacheFile -MakeMkvPath $realMkv -Now ((Get-Date).AddMinutes(6)))) "older than the 5-minute TTL: ignored"
    Set-Content -LiteralPath $cacheFile -Value $good   # an old-format cache with no header line
    Assert-True ($null -eq (Read-DriveCache -CacheFile $cacheFile -MakeMkvPath $realMkv)) "an old cache without the header line is ignored"
    Assert-True ($null -eq (Read-DriveCache -CacheFile (Join-Path $cacheDir 'none.txt') -MakeMkvPath $realMkv)) "no cache file: nothing"
    Assert-True (Test-CacheableMakeMkvPath 'C:\Program Files\MakeMKV\makemkvcon.exe') "makemkvcon.exe (32-bit name) is cacheable"
} finally {
    Remove-Item -LiteralPath $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nrip-disc.ps1 - drive lookup waits 20s per attempt, then offers a retry" -ForegroundColor Cyan
$ripText = [IO.File]::ReadAllText($ripDiscPath)
Assert-True ($ripText -match '\$drvQueryTimeoutSec = if \(\$preMkvRunning\) \{ 60 \} else \{ 20 \}') "the startup drive lookup gives up after 20s (60s only while another rip runs)"
Assert-True ($ripText -match '\$drvQueryMaxAttempts = 3\b') "up to 3 attempts"
Assert-True ($ripText -match "-TimeoutSec \`$drvQueryTimeoutSec -OnTick \`$drvTick") "the lookup shows a countdown while it waits"
Assert-True ($ripText.Contains("[R] Retry (usually quick once the drive is awake) / Enter = R / [A] Abort")) "a timeout asks to retry instead of exiting"

Write-Host "`nGet-MakeMkvProcessInfo - a running rip is never called a leftover" -ForegroundColor Cyan
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-MakeMkvProcessInfo')))
# Real command line of the 2026-11-24 Michael Collins rip (PID 36936), parent alive
$procs = @(
    [pscustomobject]@{ ProcessId = 36936; ParentProcessId = 100; CommandLine = '"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" mkv disc:1 all "C:\Video\Michael Collins (1996)\Disc1" --minlength=120 --progress=-same' },
    [pscustomobject]@{ ProcessId = 500; ParentProcessId = 999; CommandLine = '"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" -r info disc:0' }
)
$info = @(Get-MakeMkvProcessInfo -Ids @(36936, 500) -Processes $procs -AliveIds @(100))
Assert-Equal $true $info[0].Active 'parent PowerShell alive: active rip'
Assert-Equal 'saving disc:1 to C:\Video\Michael Collins (1996)\Disc1' $info[0].What 'says what the rip is doing'
Assert-Equal $false $info[1].Active 'parent gone: leftover'
Assert-Equal 'reading disc:0' $info[1].What 'info query described'
Assert-True (-not $ripText.Contains('Get-Process makemkvcon,makemkvcon64 | Stop-Process -Force')) 'no blanket "kill every MakeMKV" advice any more'
Assert-True ($ripText.Contains('That is not a leftover - do NOT close it.')) 'an active rip is called out as one to keep'
Assert-True ($ripText -match '\$preMkvRunning -and \$script:MusicLibraryChosen') 'music library offer still shown while another MakeMKV rip runs'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
