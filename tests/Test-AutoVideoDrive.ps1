#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for picking the drive that holds the DVD / Blu-ray when no -Drive is given,
    and for the "not connected" drive hint.

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 via the PowerShell AST
    parser. Disc types and PnP devices are faked; no drive is touched.

.EXAMPLE
    .\tests\Test-AutoVideoDrive.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ScriptPath has $($parseErrors.Count) parse error(s)." }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

foreach ($name in @('Get-OtherDiscDrives', 'Select-AutoVideoDrive', 'Get-DisconnectedOpticalDrives')) {
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

function New-Drv { param([int]$Index, [string]$Letter, [string]$DiscName = '', [bool]$Busy = $false)
    [pscustomobject]@{ Index = $Index; Name = "drive $Index"; DiscName = $DiscName; Letter = $Letter; Busy = $Busy } }

# Fake disc types by letter; records which letters were queried.
function New-TypeMap { param([hashtable]$Map)
    # A closure gets its own scope, so it captures the list itself rather than $script:
    $queried = [System.Collections.Generic.List[string]]::new()
    $script:Queried = $queried
    return { param($l) $queried.Add($l); "$($Map[$l])" }.GetNewClosure() }

Write-Host "`nSelect-AutoVideoDrive" -ForegroundColor Cyan

# The 2026-10-29 case: audio CD in the default D:, the DVD in G:
$lines = @((New-Drv 0 'D:' 'Audio CD'), (New-Drv 1 'G:' 'JEAN_DE_FLORETTE'))
$r = Select-AutoVideoDrive -DrvLines $lines -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video' })
Assert-Equal 'switch' $r.Action 'audio CD in the default, one DVD elsewhere: switch'
Assert-Equal 'G:' $r.Drives[0].Letter 'switches to the drive holding the DVD'
Assert-Equal 'Audio CD' $r.CurrentType 'reports what the default drive holds'

$r = Select-AutoVideoDrive -DrvLines $lines -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'DVD-Video'; 'G:' = 'Blu-ray' })
Assert-Equal 'keep' $r.Action 'the default already holds a DVD: keep it'
Assert-Equal 1 $script:Queried.Count 'only the default drive is looked at when it already holds a DVD'

$r = Select-AutoVideoDrive -DrvLines $lines -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Blu-ray'; 'G:' = 'DVD-Video' })
Assert-Equal 'keep' $r.Action 'a Blu-ray in the default counts as a video disc too'

$r = Select-AutoVideoDrive -DrvLines $lines -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = ''; 'G:' = 'Blu-ray' })
Assert-Equal 'switch' $r.Action 'empty default drive, a Blu-ray elsewhere: switch'

$three = @((New-Drv 0 'D:'), (New-Drv 1 'G:'), (New-Drv 2 'H:'))
$r = Select-AutoVideoDrive -DrvLines $three -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video'; 'H:' = 'Blu-ray' })
Assert-Equal 'ask' $r.Action 'two other drives with video discs: ask which'
Assert-Equal 2 $r.Drives.Count 'both are offered'

$r = Select-AutoVideoDrive -DrvLines $three -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD'; 'G:' = 'Audio CD'; 'H:' = 'Data disc (DVD-sized)' })
Assert-Equal 'keep' $r.Action 'no video disc anywhere: keep the default (the audio-CD guard takes over)'

$r = Select-AutoVideoDrive -DrvLines @((New-Drv 0 'D:')) -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD' })
Assert-Equal 'keep' $r.Action 'only one drive connected: nothing to switch to'

$busy = @((New-Drv 0 'D:' 'OTHER FILM' $true), (New-Drv 1 'G:'))
$r = Select-AutoVideoDrive -DrvLines $busy -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'DVD-Video'; 'G:' = 'DVD-Video' })
Assert-Equal 'switch' $r.Action 'the default is busy with another rip: use the free drive with a DVD'
Assert-Equal $true $r.CurrentBusy 'reports the default as busy'
Assert-Equal $false ($script:Queried -contains 'D:') 'a busy drive is never queried'

$busyOther = @((New-Drv 0 'D:'), (New-Drv 1 'G:' 'X' $true))
$r = Select-AutoVideoDrive -DrvLines $busyOther -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video' })
Assert-Equal 'keep' $r.Action 'a DVD only in a busy drive is not taken'

$r = Select-AutoVideoDrive -DrvLines @((New-Drv 1 'G:')) -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'D:' = ''; 'G:' = 'DVD-Video' })
Assert-Equal 'switch' $r.Action 'the default is not in MakeMKV''s list at all: use the drive with the DVD'

$r = Select-AutoVideoDrive -DrvLines $lines -CurrentLetter 'd' -GetDiscType (New-TypeMap @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video' })
Assert-Equal 'G:' $r.Drives[0].Letter 'a bare lower-case letter for the default still works'

Write-Host "`nGet-DisconnectedOpticalDrives" -ForegroundColor Cyan
$dev = { param($n, $p) [pscustomobject]@{ FriendlyName = "$n USB Device"; Present = $p } }
# Tonight's PnP list: the BU40N and the 1.02 GP75N not present, the 1.01 GP75N connected
$devices = @((& $dev 'HL-DT-ST BD-RE BU40N' $false), (& $dev 'HL-DT-ST DVDRAM GP75N' $true), (& $dev 'HL-DT-ST DVDRAM GP75N' $false))
$missing = @(Get-DisconnectedOpticalDrives -Devices $devices)
Assert-Equal 2 $missing.Count 'two drives are not connected'
Assert-Equal 'HL-DT-ST BD-RE BU40N' $missing[0] 'the missing BU40N is named, without " USB Device"'
Assert-Equal 'HL-DT-ST DVDRAM GP75N (another one)' $missing[1] 'a missing twin of a connected drive says "another one"'
Assert-Equal 0 @(Get-DisconnectedOpticalDrives -Devices @((& $dev 'HL-DT-ST BD-RE BU40N' $true))).Count 'all connected: nothing listed'

Write-Host "`n$script:Passed passed, $script:Failed failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
