#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for watch-drives.ps1 (insert a disc -> rip-disc.ps1 opens in a new window).

.DESCRIPTION
    Runs against the real function bodies lifted out of watch-drives.ps1 at run time via the
    PowerShell AST parser, so the tests cannot drift from the shipped code without failing.

    The decisions (which drive, which disc kind, duplicate events, busy drive) are pure and
    tested against fixtures and temp folders. The new window's side (-RunRipFor) is run for
    real, in this process, against a fake rip-disc.ps1 - including one that calls 'exit' and
    one that throws - to prove the arguments arrive and the lock file is always removed.

    Deliberately NO child powershell.exe is started here. The first version of this file
    (tests\Test-WatchDrives.ps1) started a windowless one with a base64 command, and
    antivirus (AVG behaviour shield) deleted that file mid-run and then refused to let the
    name be created again, 2026-10-11. Hence this file's different name.

    Not covered here: the WMI insert event itself and the visible window - those need a
    real disc going into a real drive.

.EXAMPLE
    .\tests\Test-DriveWatcher.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$watchPath = Join-Path $repoRoot 'watch-drives.ps1'

if (-not (Test-Path $watchPath)) { throw "Cannot find $watchPath - run this from inside the repo." }

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

foreach ($name in 'Get-NormalizedDriveLetter', 'Test-WatchedDrive', 'Test-ArrivalIsDuplicate', 'Get-DiscKind',
                  'Get-InsertAction', 'Test-RipBusy', 'Get-RipCommandText', 'Get-RipWindowArguments') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $watchPath -FunctionName $name)))
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
    Assert-Equal $true $Condition $Because
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("watch-drives-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    Write-Host "`nTest-WatchedDrive" -ForegroundColor Cyan

    Assert-True (Test-WatchedDrive -DriveLetter 'D:' -Watch @()) "an empty -Drive list watches every drive"
    Assert-True (Test-WatchedDrive -DriveLetter 'E:' -Watch @('D:', 'E:')) "a listed drive is watched"
    Assert-Equal $false (Test-WatchedDrive -DriveLetter 'F:' -Watch @('D:', 'E:')) "an unlisted drive is not watched"
    Assert-True (Test-WatchedDrive -DriveLetter 'd:' -Watch @('D:')) "drive letters match regardless of case"
    Assert-Equal 'D:' (Get-NormalizedDriveLetter 'D:\') "a -Drive value typed as D:\ is normalised to D:"
    Assert-Equal 'D:' (Get-NormalizedDriveLetter 'D') "a -Drive value typed as D is normalised to D:"

    Write-Host "`nTest-ArrivalIsDuplicate" -ForegroundColor Cyan

    $seen = @{}
    $t0 = [datetime]'2026-10-11T10:00:00'
    Assert-Equal $false (Test-ArrivalIsDuplicate -LastSeen $seen -DriveLetter 'D:' -Now $t0 -DebounceSeconds 10) "the first event for a drive is not a duplicate"
    Assert-True (Test-ArrivalIsDuplicate -LastSeen $seen -DriveLetter 'D:' -Now $t0.AddSeconds(3) -DebounceSeconds 10) "a second event 3s later is a duplicate"
    Assert-Equal $false (Test-ArrivalIsDuplicate -LastSeen $seen -DriveLetter 'E:' -Now $t0.AddSeconds(4) -DebounceSeconds 10) "another drive's event in the same window is not a duplicate"
    Assert-Equal $false (Test-ArrivalIsDuplicate -LastSeen $seen -DriveLetter 'D:' -Now $t0.AddSeconds(60) -DebounceSeconds 10) "the next disc a minute later is not a duplicate"

    Write-Host "`nGet-DiscKind" -ForegroundColor Cyan

    $dvd = Join-Path $tempRoot 'dvd';   New-Item -ItemType Directory -Path (Join-Path $dvd 'VIDEO_TS') -Force | Out-Null
    $bd = Join-Path $tempRoot 'bd';     New-Item -ItemType Directory -Path (Join-Path $bd 'BDMV') -Force | Out-Null
    $data = Join-Path $tempRoot 'data'; New-Item -ItemType Directory -Path $data -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $data 'readme.txt') -Value 'x'
    $cd = Join-Path $tempRoot 'cd';     New-Item -ItemType Directory -Path $cd -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $cd 'Track01.cda') -Value 'x'

    Assert-Equal 'DVD-Video' (Get-DiscKind -Root $dvd -VolumeLabel 'SILICON_VALLEY_S2_D2') "a VIDEO_TS folder means DVD-Video"
    Assert-Equal 'Blu-ray' (Get-DiscKind -Root $bd -VolumeLabel 'INCEPTION') "a BDMV folder means Blu-ray"
    Assert-Equal 'Data disc' (Get-DiscKind -Root $data -VolumeLabel 'BURN') "a readable disc with neither folder is a data disc"
    Assert-Equal 'Audio CD' (Get-DiscKind -Root $data -VolumeLabel 'Audio CD') "the 'Audio CD' volume label means an audio CD"
    Assert-Equal 'Audio CD' (Get-DiscKind -Root $cd -VolumeLabel '') ".cda tracks mean an audio CD even without the label"
    Assert-Equal '' (Get-DiscKind -Root (Join-Path $tempRoot 'missing') -VolumeLabel '') "an unreadable root gives '' rather than a guess"

    Write-Host "`nGet-InsertAction" -ForegroundColor Cyan

    Assert-True (Get-InsertAction -DiscKind 'DVD-Video').Rip "a DVD-Video starts a rip"
    Assert-True (Get-InsertAction -DiscKind 'Blu-ray').Rip "a Blu-ray starts a rip"
    $audio = Get-InsertAction -DiscKind 'Audio CD'
    Assert-Equal $false $audio.Rip "an audio CD does not start a rip"
    Assert-True ($audio.Reason -like '*rip-audio.ps1*') "the audio CD reason points at ripaudio's rip-audio.ps1"
    Assert-Equal $false (Get-InsertAction -DiscKind 'Data disc').Rip "a data disc does not start a rip"
    Assert-Equal $false (Get-InsertAction -DiscKind '').Rip "an unreadable disc does not start a rip"

    Write-Host "`nTest-RipBusy" -ForegroundColor Cyan

    Assert-True (Test-RipBusy -ProcessAlive $true -LockExists $true) "window open and rip not finished: busy"
    Assert-Equal $false (Test-RipBusy -ProcessAlive $true -LockExists $false) "window left open after the rip finished: not busy"
    Assert-Equal $false (Test-RipBusy -ProcessAlive $false -LockExists $true) "window closed mid-rip (stale lock): not busy"

    Write-Host "`nGet-RipCommandText / Get-RipWindowArguments" -ForegroundColor Cyan

    Assert-Equal "& 'C:\RipDisc\rip-disc.ps1' -Drive E: -OutputDrive F: -NoSound" `
        (Get-RipCommandText -RipScriptPath 'C:\RipDisc\rip-disc.ps1' -DriveLetter 'E:' -RipArguments ' -OutputDrive F: -NoSound ') `
        "the command runs rip-disc.ps1 with -Drive and the extra arguments"
    Assert-Equal "& 'C:\RipDisc\rip-disc.ps1' -Drive D:" `
        (Get-RipCommandText -RipScriptPath 'C:\RipDisc\rip-disc.ps1' -DriveLetter 'D:') "no extra arguments leaves just -Drive"
    Assert-Equal "& 'C:\RipDisc\rip-disc.ps1' -Drive D:" `
        (Get-RipCommandText -RipScriptPath 'C:\RipDisc\rip-disc.ps1' -DriveLetter 'D:' -RipArguments "`r`n") "a blank lock file (no extra arguments) adds nothing"
    Assert-True ((Get-RipCommandText -RipScriptPath "C:\Steve's Tools\rip-disc.ps1" -DriveLetter 'D:') -like "*'C:\Steve''s Tools\rip-disc.ps1'*") `
        "an apostrophe in the script path is escaped"

    $windowArgs = Get-RipWindowArguments -WatchScriptPath 'C:\Rip Disc\watch-drives.ps1' -DriveLetter 'E:' -LockPath 'C:\Temp\ripdisc-watch-E.lock'
    Assert-Equal '-NoExit -File "C:\Rip Disc\watch-drives.ps1" -RunRipFor E: -LockPath "C:\Temp\ripdisc-watch-E.lock"' $windowArgs `
        "the new window re-runs watch-drives.ps1 in -RunRipFor mode with plain arguments only, paths quoted"

    Write-Host "`n-RunRipFor, run for real against a fake rip-disc.ps1" -ForegroundColor Cyan

    function Invoke-FakeRip {
        param([string]$FakeBody, [string]$RipArguments = "")

        $dir = Join-Path $tempRoot ("it's a run " + [guid]::NewGuid().ToString('N').Substring(0, 6))  # space + apostrophe on purpose
        New-Item -ItemType Directory -Path $dir | Out-Null
        $fake = Join-Path $dir 'rip-disc.ps1'
        $out = Join-Path $dir 'args.txt'
        $lock = Join-Path $dir 'drive.lock'
        Set-Content -LiteralPath $fake -Value ("param([string]`$Drive, [string]`$OutputDrive, [switch]`$NoSound)`n" +
            "Set-Content -LiteralPath '$($out.Replace("'", "''"))' -Value `"`$Drive|`$OutputDrive|`$NoSound`"`n$FakeBody")
        Set-Content -LiteralPath $lock -Value $RipArguments   # exactly how the watcher hands the arguments over

        $threw = $false
        try { & $watchPath -RunRipFor 'E' -RipScriptPath $fake -LockPath $lock } catch { $threw = $true }
        return @{
            Args = if (Test-Path -LiteralPath $out) { (Get-Content -LiteralPath $out -Raw).Trim() } else { '<fake script never ran>' }
            LockRemoved = -not (Test-Path -LiteralPath $lock)
            Threw = $threw
        }
    }

    $ok = Invoke-FakeRip -FakeBody '' -RipArguments '-OutputDrive F: -NoSound'
    Assert-Equal 'E:|F:|True' $ok.Args "rip-disc.ps1 receives -Drive and the extra arguments (path with a space and an apostrophe)"
    Assert-True $ok.LockRemoved "the lock file is removed when rip-disc.ps1 finishes"

    $bare = Invoke-FakeRip -FakeBody ''
    Assert-Equal 'E:||False' $bare.Args "with no extra arguments rip-disc.ps1 receives just -Drive"

    $exited = Invoke-FakeRip -FakeBody 'exit 1'
    Assert-Equal 'E:||False' $exited.Args "the failing fake script ran"
    Assert-True $exited.LockRemoved "the lock file is removed when rip-disc.ps1 calls 'exit' (a failed rip)"
    Assert-True $true "'exit' inside rip-disc.ps1 did not end the calling session (this line was reached)"

    $boom = Invoke-FakeRip -FakeBody 'throw "boom"'
    Assert-True $boom.Threw "an error thrown by rip-disc.ps1 is not swallowed"
    Assert-True $boom.LockRemoved "the lock file is removed when rip-disc.ps1 throws"
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
