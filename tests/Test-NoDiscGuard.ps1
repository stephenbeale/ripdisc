#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the no-disc guard: an empty drive must stop before auto-discovery.

.DESCRIPTION
    2026-11-27: D: was empty, MakeMKV still listed the drive, its scan found 0 titles and
    auto-discovery offered title searches. Get-DriveDiscState (lifted from rip-disc.ps1 via
    the AST) is run with faked disc types and media flags; no drive is touched. The guard's
    position in rip-disc.ps1 is checked too.

.EXAMPLE
    .\tests\Test-NoDiscGuard.ps1
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
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-DriveDiscState')))

$script:Passed = 0
$script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS: $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL: $Because (expected '$Expected', got '$Actual')" -ForegroundColor Red }
}

$noSleep = { }
$quiet = { param($l, $s) $script:Waited = $true }

Write-Host "`nGet-DriveDiscState" -ForegroundColor Cyan

$script:Waited = $false
$r = Get-DriveDiscState -DriveLetter 'D:' -GetDiscType { param($l) 'DVD-Video' } -GetMediaLoaded { param($l) $true } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'ok' $r 'a DVD is ok'
Assert-Equal $false $script:Waited 'no wait when the disc reads at once'

$script:Waited = $false
$r = Get-DriveDiscState -DriveLetter 'D:' -GetDiscType { param($l) '' } -GetMediaLoaded { param($l) $false } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'empty' $r 'no media (the 2026-11-27 case) is empty'
Assert-Equal $false $script:Waited 'an empty drive is not waited on'

$r = Get-DriveDiscState -DriveLetter 'D:' -GetDiscType { param($l) 'Data disc (DVD-ROM-sized)' } -GetMediaLoaded { param($l) $true } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'ok' $r 'a readable data disc is left to the later checks'

# Disc still spinning up: unreadable for 3 checks, then a DVD
$script:Calls = 0
$spinUp = { param($l) $script:Calls++; if ($script:Calls -gt 3) { 'DVD-Video' } else { '' } }
$script:Waited = $false
$r = Get-DriveDiscState -DriveLetter 'D:' -WaitSec 10 -GetDiscType $spinUp -GetMediaLoaded { param($l) $null } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'ok' $r 'a disc that reads after a few seconds is ok'
Assert-Equal $true $script:Waited 'waits while the disc spins up'
Assert-Equal 4 $script:Calls 'stops checking as soon as it reads'

$script:Calls = 0
$never = { param($l) $script:Calls++; '' }
$r = Get-DriveDiscState -DriveLetter 'D:' -WaitSec 5 -GetDiscType $never -GetMediaLoaded { param($l) $true } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'unreadable' $r 'media loaded but never readable is unreadable'
Assert-Equal 6 $script:Calls 'one check up front plus one per second of the wait'

$r = Get-DriveDiscState -DriveLetter 'D:' -WaitSec 3 -GetDiscType { param($l) '' } -GetMediaLoaded { param($l) $null } -Sleep $noSleep -OnWait $quiet
Assert-Equal 'unreadable' $r 'unknown media flag and unreadable is unreadable (offers carry on)'

# Media flag unknown at first, then the tray turns out empty
$script:MediaCalls = 0
$laterEmpty = { param($l) $script:MediaCalls++; if ($script:MediaCalls -eq 1) { $null } else { $false } }
$r = Get-DriveDiscState -DriveLetter 'D:' -WaitSec 2 -GetDiscType { param($l) '' } -GetMediaLoaded $laterEmpty -Sleep $noSleep -OnWait $quiet
Assert-Equal 'empty' $r 'media flag read again after the wait'

Write-Host "`nGuard placement in rip-disc.ps1" -ForegroundColor Cyan
$src = Get-Content -LiteralPath $ripDiscPath -Raw
$guard = $src.IndexOf('# ========== NO DISC GUARD ==========')
$audio = $src.IndexOf('# ========== AUDIO CD GUARD ==========')
$disc  = $src.IndexOf('AUTO-DISCOVERY MODE')
Assert-Equal $true ($guard -gt 0) 'guard present'
Assert-Equal $true ($guard -gt $audio) 'after the audio CD guard (an audio CD gets the ripaudio offer, not "no disc")'
Assert-Equal $true ($guard -lt $disc) 'before auto-discovery'

Write-Host "`nAnything to rip? (before the first question)" -ForegroundColor Cyan
$loadConfig = Join-Path $repoRoot 'Load-Config.ps1'
foreach ($n in 'Get-FreeDiscSummary', 'Get-NothingToRipMessage') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $loadConfig -FunctionName $n)))
}
$d = { param($letter, $loaded) [pscustomobject]@{ Drive = $letter; MediaLoaded = $loaded } }

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $false), (& $d 'G:' $false)) -MkvRipCount 0 -AudioBusyLetters @()
Assert-Equal 'No disc in any drive - nothing to rip.' (Get-NothingToRipMessage $s) 'every drive empty (first 2026-11-27 run)'

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $true)) -MkvRipCount 1 -AudioBusyLetters @()
Assert-Equal 'Every disc in a drive (D:) is already being ripped (1 rip running) - nothing new to rip.' (Get-NothingToRipMessage $s) 'only disc is being ripped by MakeMKV (second 2026-11-27 run)'

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $true), (& $d 'G:' $true)) -MkvRipCount 1 -AudioBusyLetters @()
Assert-Equal '' (Get-NothingToRipMessage $s) 'a second disc while one rip runs: carry on'
Assert-Equal 1 $s.Free 'one free disc'

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $true), (& $d 'G:' $true)) -MkvRipCount 1 -AudioBusyLetters @('g:')
Assert-Equal 'Every disc in a drive (D:, G:) is already being ripped (2 rips running) - nothing new to rip.' (Get-NothingToRipMessage $s) 'DVD rip on one drive, ripaudio on the other'

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $true)) -MkvRipCount 0 -AudioBusyLetters @('F:')
Assert-Equal '' (Get-NothingToRipMessage $s) 'audio rip on a drive with no disc listed does not hide a free disc'

$s = Get-FreeDiscSummary -Drives @() -MkvRipCount 0 -AudioBusyLetters @()
Assert-Equal $false $s.Known 'no optical drives seen: unknown'
Assert-Equal '' (Get-NothingToRipMessage $s) 'unknown never stops the run'

$s = Get-FreeDiscSummary -Drives @((& $d 'D:' $true)) -MkvRipCount 3 -AudioBusyLetters @()
Assert-Equal 0 $s.Free 'free never below 0'

$opIdx = $src.IndexOf('Read-OutputDrive -Default')
$earlyIdx = $src.IndexOf('# ========== ANYTHING TO RIP? ==========')
Assert-Equal $true ($earlyIdx -gt 0 -and $earlyIdx -lt $opIdx) 'early check runs before the output drive question'

Write-Host "`n$script:Passed passed, $script:Failed failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
