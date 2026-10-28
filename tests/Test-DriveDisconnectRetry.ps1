#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for Step 1's pause-and-retry when the optical drive disconnects mid-rip.

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 at run time via the
    PowerShell AST parser, so the tests cannot drift from the shipped code without failing.

    Covers detecting MakeMKV's disconnect signature, finding a replugged drive's (possibly new)
    disc:N index - both from fixture output and through a fake makemkvcon - and the
    Retry/Abort prompt, with Read-Host replaced by a scripted answer queue and the
    "is there a disc in the drive" check replaced by a stub.

    The retry loop itself is inline in the Step 1 body and needs a real drive to be unplugged
    to exercise end to end; that is a manual test.

.EXAMPLE
    .\tests\Test-DriveDisconnectRetry.ps1
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

foreach ($name in 'Test-MakeMkvDriveDisconnected', 'Get-MakeMkvDriveIndexFromOutput', 'Find-MakeMkvDriveIndex', 'Wait-ProcessWithTimeout', 'Wait-DriveReconnect') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $name)))
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

Write-Host "`nTest-MakeMkvDriveDisconnected" -ForegroundColor Cyan

# The live failure that prompted this feature (Cream documentary, knocked USB cable).
$disconnectOutput = @(
    'Saving 21 titles into directory file://C:/Video/Cream/Disc1'
    'Error ''Posix error - STATUS_DEVICE_NOT_CONNECTED'' occurred while reading ''DVD+R-DL HL-DT-ST DVDRAM GP75N'' at offset ''1048576'''
    'Copy complete. 0 titles saved, 21 failed.'
) -join "`n"
Assert-Equal $true (Test-MakeMkvDriveDisconnected -OutputText $disconnectOutput) "STATUS_DEVICE_NOT_CONNECTED is a disconnect"
Assert-Equal $true (Test-MakeMkvDriveDisconnected -OutputText "Failed to open a device which does not exist") "'device which does not exist' is a disconnect"
Assert-Equal $true (Test-MakeMkvDriveDisconnected -OutputText "the device does not exist") "'device does not exist' is a disconnect"
Assert-Equal $false (Test-MakeMkvDriveDisconnected -OutputText "Scsi error - SCRAMBLED SECTOR WITHOUT AUTHENTICATION") "a CSS failure is not a disconnect (must not be retried)"
Assert-Equal $false (Test-MakeMkvDriveDisconnected -OutputText "Copy complete. 3 titles saved.") "a clean rip is not a disconnect"
Assert-Equal $false (Test-MakeMkvDriveDisconnected -OutputText "") "empty output is not a disconnect"

Write-Host "`nGet-MakeMkvDriveIndexFromOutput" -ForegroundColor Cyan

$drvOutput = @(
    'MSG:1005,0,1,"MakeMKV v1.18.4 win(x64-release) started","%1 started","MakeMKV v1.18.4 win(x64-release)"'
    'DRV:0,2,999,1,"BD-RE HL-DT-ST BD-RE BU40N 1.05 MO4P6N95940","SILICON VALLEY S1 D2","D:"'
    'DRV:1,256,999,0,"","",""'
    'DRV:2,2,999,1,"DVD+R-DL HL-DT-ST DVDRAM GP75N 1.02 K0GPAID0457","CREAM","E:"'
    'DRV:3,256,999,0,"","",""'
)
Assert-Equal 2 (Get-MakeMkvDriveIndexFromOutput -Lines $drvOutput -DriveLetter "E:") "finds the drive letter's index"
Assert-Equal 0 (Get-MakeMkvDriveIndexFromOutput -Lines $drvOutput -DriveLetter "D:") "finds index 0"
Assert-Equal -1 (Get-MakeMkvDriveIndexFromOutput -Lines $drvOutput -DriveLetter "F:") "returns -1 when the letter is not listed (still unplugged)"
Assert-Equal -1 (Get-MakeMkvDriveIndexFromOutput -Lines @('DRV:1,256,999,0,"","","E:"') -DriveLetter "E:") "skips empty slots (flag 256)"

# A replugged USB drive can come back under a different index.
$replugged = @(
    'DRV:0,2,999,1,"BD-RE HL-DT-ST BD-RE BU40N 1.05 MO4P6N95940","","D:"'
    'DRV:3,2,999,1,"DVD+R-DL HL-DT-ST DVDRAM GP75N 1.02 K0GPAID0457","CREAM","E:"'
)
Assert-Equal 3 (Get-MakeMkvDriveIndexFromOutput -Lines $replugged -DriveLetter "E:") "picks up the new index after a replug"
Assert-Equal -1 (Get-MakeMkvDriveIndexFromOutput -Lines @() -DriveLetter "E:") "empty output returns -1"

Write-Host "`nFind-MakeMkvDriveIndex - against a fake makemkvcon" -ForegroundColor Cyan

$fakeDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-disconnect-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fakeDir | Out-Null
try {
    $fakeMkv = Join-Path $fakeDir 'fake-makemkvcon.cmd'
    Set-Content -Path $fakeMkv -Encoding ASCII -Value @(
        '@echo off'
        'echo DRV:0,2,999,1,"BD-RE DRIVE","","D:"'
        'echo DRV:4,2,999,1,"DVD DRIVE","CREAM","E:"'
    )
    Assert-Equal 4 (Find-MakeMkvDriveIndex -MakeMkvPath $fakeMkv -DriveLetter "E:" -TimeoutSec 15) "reads the index from MakeMKV's live output"
    Assert-Equal -1 (Find-MakeMkvDriveIndex -MakeMkvPath $fakeMkv -DriveLetter "G:" -TimeoutSec 15) "returns -1 when MakeMKV does not list the drive"
    Assert-Equal -1 (Find-MakeMkvDriveIndex -MakeMkvPath (Join-Path $fakeDir 'missing.exe') -DriveLetter "E:" -TimeoutSec 15) "returns -1 (no throw) when makemkvcon cannot be started"
} finally {
    Remove-Item -LiteralPath $fakeDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nWait-DriveReconnect - Retry/Abort prompt" -ForegroundColor Cyan

# Read-Host is replaced by a function in this scope, which PowerShell resolves before the
# cmdlet, so Wait-DriveReconnect reads scripted answers. $null in the queue simulates end of input.
$script:Answers = New-Object System.Collections.Queue
$script:Prompts = 0
function Read-Host {
    param([string]$Prompt)
    $script:Prompts++
    if ($script:Answers.Count -eq 0) { return $null }
    return $script:Answers.Dequeue()
}
function Set-Answers { param([object[]]$Values) $script:Answers.Clear(); foreach ($v in $Values) { $script:Answers.Enqueue($v) }; $script:Prompts = 0 }

$ready = { param($Root) $true }
$notReady = { param($Root) $false }

Set-Answers @('R')
Assert-Equal $true (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $ready 6>$null) "R with the disc readable retries"

Set-Answers @('r')
Assert-Equal $true (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $ready 6>$null) "lower-case r is accepted"

Set-Answers @('A')
Assert-Equal $false (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $ready 6>$null) "A aborts"

Set-Answers @()
Assert-Equal $false (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $ready 6>$null) "end of input aborts rather than retrying forever"

Set-Answers @('R', 'R', 'A')
Assert-Equal $false (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $notReady 6>$null) "R is refused while the drive still has no readable disc"
Assert-Equal 3 $script:Prompts "asks again after each refused R"

# Drive comes back on the second check.
$script:ReadyChecks = 0
$comesBack = { param($Root) $script:ReadyChecks++; $script:ReadyChecks -ge 2 }
Set-Answers @('R', 'R')
Assert-Equal $true (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $comesBack 6>$null) "retries once the drive reappears"
Assert-Equal 2 $script:Prompts "only one extra prompt was needed"

$script:CheckedRoot = ""
$recordRoot = { param($Root) $script:CheckedRoot = $Root; $true }
Set-Answers @('R')
[void](Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $recordRoot 6>$null)
Assert-Equal "E:\" $script:CheckedRoot "checks the drive root (E:\)"

Set-Answers @('R')
Assert-Equal $true (Wait-DriveReconnect -DriveLetter "D:" -CheckDriveLetter $false -IsMediaReady $notReady 6>$null) "with -DriveIndex (no reliable letter) R retries without the media check"

Set-Answers @('x', '', 'A')
Assert-Equal $false (Wait-DriveReconnect -DriveLetter "E:" -IsMediaReady $ready 6>$null) "other input re-prompts instead of choosing for the user"
Assert-Equal 3 $script:Prompts "re-prompted for each invalid answer"

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
