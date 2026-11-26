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

foreach ($name in @('Get-OtherDiscDrives', 'Select-AutoVideoDrive', 'Get-DisconnectedOpticalDrives', 'Get-WindowsOpticalDrives', 'Get-AudioCdShortcutDrive', 'Get-CyanripBusyDrives')) {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $name)))
}
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath (Join-Path $repoRoot 'Load-Config.ps1') -FunctionName 'Test-MusicLibraryFolder')))

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

Write-Host "`nGet-AudioCdShortcutDrive (before MakeMKV is asked anything)" -ForegroundColor Cyan
# 2026-11-22: audio CDs in both D: and G:, no -Drive - MakeMKV's lookup timed out at 60 s
Assert-Equal 'D:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = 'Audio CD'; 'G:' = 'Audio CD' }) 'two audio CDs, no -Drive: hand the default D: to ripaudio without MakeMKV'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = 'Audio CD'; 'G:' = 'DVD-Video' }) 'no -Drive and a DVD in G: - carry on, the auto drive pick takes G:'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = 'Audio CD'; 'G:' = 'Blu-ray' }) 'no -Drive and a Blu-ray in G: - carry on'
Assert-Equal 'G:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = ''; 'G:' = 'Audio CD' }) 'no -Drive, default empty, only an audio CD elsewhere: hand that one over'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = ''; 'G:' = '' }) 'no discs at all: carry on (the normal path reports it)'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{ 'D:' = 'DVD-Video'; 'G:' = 'Audio CD' }) 'no -Drive, DVD in the default: carry on'
Assert-Equal 'G:' (Get-AudioCdShortcutDrive -TargetLetter 'G:' -DriveGiven -DiscTypes @{ 'D:' = 'DVD-Video'; 'G:' = 'Audio CD' }) '-Drive G with an audio CD: hand over even though D: has a DVD (guard offers [D])'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DriveGiven -DiscTypes @{ 'D:' = 'DVD-Video'; 'G:' = 'Audio CD' }) '-Drive D with a DVD: carry on'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DriveGiven -DiscTypes @{ 'D:' = ''; 'G:' = 'Audio CD' }) '-Drive D (empty) is never second-guessed'
Assert-Equal 'D:' (Get-AudioCdShortcutDrive -TargetLetter 'd' -DiscTypes @{ 'D:' = 'Audio CD' }) 'a bare lower-case letter works'
Assert-Equal '' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes @{}) 'Windows listed no drives: carry on'

Write-Host "`nGet-WindowsOpticalDrives" -ForegroundColor Cyan
$cd = { param($drive, $loaded, $vol) [pscustomobject]@{ Drive = $drive; Name = 'HL-DT-ST DVDRAM GP75N USB Device'; MediaLoaded = $loaded; VolumeName = $vol } }
$w = @(Get-WindowsOpticalDrives -Devices @((& $cd 'G:' $true 'Audio CD'), (& $cd 'D:' $true 'Audio CD'), (& $cd 'E:' $false 'OLD'), (& $cd '' $false '')))
Assert-Equal 3 $w.Count 'rows for each lettered drive, none for a drive without a letter'
Assert-Equal 'D:' $w[0].Letter 'sorted by letter'
Assert-Equal 'Audio CD' $w[0].DiscName 'volume name used as the disc name'
Assert-Equal '' $w[1].DiscName 'no media: no disc name, even if WMI kept an old label'
Assert-Equal $false $w[0].Busy 'never marked busy (the check is skipped while MakeMKV runs)'
$others = @(Get-OtherDiscDrives -DrvLines $w -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'G:' = 'Audio CD' }))
Assert-Equal 'G:' $others[0].Letter 'Windows rows work with Get-OtherDiscDrives (the guard''s [D] / other-audio-CD lists)'

Write-Host "`nTest-MusicLibraryFolder" -ForegroundColor Cyan
Assert-Equal $true (Test-MusicLibraryFolder 'Music') 'Music'
Assert-Equal $true (Test-MusicLibraryFolder 'Music Videos') 'Music Videos'
Assert-Equal $true (Test-MusicLibraryFolder 'Live music') 'Live music (any case)'
Assert-Equal $true (Test-MusicLibraryFolder 'F:\Music\') 'a full path: the folder name counts'
Assert-Equal $false (Test-MusicLibraryFolder 'Documentaries') 'Documentaries'
Assert-Equal $false (Test-MusicLibraryFolder 'F:\Music\Concerts') 'only the library folder itself, not a parent'
Assert-Equal $false (Test-MusicLibraryFolder '') 'no folder (the Movies row)'

Write-Host "`nrip-disc.ps1 order" -ForegroundColor Cyan
$ripText = [IO.File]::ReadAllText($ripDiscPath)
$preAt = $ripText.IndexOf('# ========== AUDIO CD CHECK BEFORE MAKEMKV ==========')
$musicAt = $ripText.IndexOf('# ========== MUSIC LIBRARY OFFER ==========')
$queryAt = $ripText.IndexOf('$drvProc.StartInfo.Arguments = "-r info disc:9999"')
Assert-Equal $true ($preAt -gt 0 -and $preAt -lt $queryAt) 'the audio CD check runs before the MakeMKV drive query'
Assert-Equal $true ($musicAt -gt $preAt -and $musicAt -lt $queryAt) 'the music library offer runs before the MakeMKV drive query'
Assert-Equal $true ($ripText.IndexOf('$script:MusicLibraryChosen = Test-MusicLibraryFolder') -gt 0) 'the library pick sets the music flag'

Write-Host "`nBusy drives (ripaudio already ripping)" -ForegroundColor Cyan
# The real command line seen 2026-11-22 while ripaudio ripped D:
$live = '"C:\Users\x\cyanrip.exe" -D "Tom Jones and Jools Holland" -o flac -d D: -s 0 -N'
Assert-Equal 'D:' ((Get-CyanripBusyDrives -CommandLines @($live)) -join ',') 'the drive comes from -d, not the -D output folder'
Assert-Equal 'D:,G:' ((Get-CyanripBusyDrives -CommandLines @($live, 'cyanrip.exe -d g -s 0')) -join ',') 'two rips; a bare letter without the colon'
Assert-Equal '' ((Get-CyanripBusyDrives -CommandLines @('cyanrip.exe -D "Disc D" -o flac')) -join ',') 'only -D (output folder): nothing busy'
Assert-Equal 0 @(Get-CyanripBusyDrives -CommandLines @()).Count 'no cyanrip running: nothing busy'
$two = @{ 'D:' = 'Audio CD'; 'G:' = 'Audio CD' }
Assert-Equal 'G:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes $two -BusyLetters @('D:')) 'default D: busy ripping: hand over the free G: (the 2026-11-22 case)'
Assert-Equal 'D:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes $two -BusyLetters @('G:')) 'G: busy: the default D: is still used'
Assert-Equal 'D:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DiscTypes $two -BusyLetters @('D:', 'G:')) 'every audio CD drive busy: the default (the guard then says why)'
Assert-Equal 'D:' (Get-AudioCdShortcutDrive -TargetLetter 'D:' -DriveGiven -DiscTypes $two -BusyLetters @('D:')) '-Drive D is never second-guessed, even when busy'
$ripText2 = [IO.File]::ReadAllText($ripDiscPath)
Assert-Equal $true ($ripText2.Contains('-DiscTypes $winTypes -BusyLetters $cyBusy')) 'the pre-check passes the busy drives'
Assert-Equal $true ($ripText2.Contains('$freeAudio = @($audioDrives | Where-Object { -not $_.Busy })')) 'the music offer only offers free drives'

# 2026-11-25: no drive is touched before it is known to be free. The default drive used to be
# woken (Test-Path D:\) first thing, while ripaudio was ripping D:.
$wakeLine = '$null = Test-Path "${driveLetter}\" -ErrorAction SilentlyContinue'
$autoStart = $ripText2.IndexOf('# ========== AUTO-DISCOVERY ==========')
$preCheck = $ripText2.IndexOf('# ========== AUDIO CD CHECK BEFORE MAKEMKV ==========')
$firstWake = $ripText2.IndexOf($wakeLine, $autoStart)
Assert-Equal $true ($autoStart -gt 0 -and $preCheck -gt $autoStart) 'auto-discovery and pre-check markers found'
Assert-Equal $true ($firstWake -gt $preCheck) 'the target drive is not woken before the audio CD / busy checks'
Assert-Equal $true ($ripText2.Contains('if ($DriveIndex -lt 0 -and -not $preMkvRunning -and $cyBusy -notcontains $driveLetter) {')) 'the wake skips a drive ripaudio is ripping, and any drive while another MakeMKV runs'
Assert-Equal $true ($ripText2.Contains('if ($w.Busy) { ''Audio CD'' } else { Get-DiscTypeLabel -DriveLetter $w.Letter }')) 'a busy drive is never probed for its disc type'

# 2026-11-26: King George ripping D: (disc:1), Meek's Cutoff in F:, no -Drive. MakeMKV's
# drive query timed out behind the running rip. Now the busy drive is found without MakeMKV
# and the run moves on to the free drive, with its index worked out.
Write-Host "`nAnother MakeMKV rip running: next free drive" -ForegroundColor Cyan
foreach ($name in @('Get-MakeMkvRips', 'Get-MakeMkvBusyDrives', 'Get-FreeDriveIndex')) {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $name)))
}
$kgCmd = '"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" mkv disc:1 all "C:\Video\The Madness of King George\Disc1" --minlength=120 --progress=-same'
$rips = @(Get-MakeMkvRips -CommandLines @($kgCmd, '"makemkvcon64.exe" -r info disc:9999'))
Assert-Equal 1 $rips.Count 'only the mkv rip counts, not an info query'
Assert-Equal 1 $rips[0].Index 'rip index read from the command line'
Assert-Equal 'C:\Video\The Madness of King George\Disc1' $rips[0].OutDir 'rip folder read from the command line'

$tmpLogs = Join-Path ([IO.Path]::GetTempPath()) ("ripdisc-busy-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpLogs | Out-Null
try {
    $kgLog = Join-Path $tmpLogs 'kg.log'
    Set-Content -LiteralPath $kgLog -Value @(
        '[2026-11-26 20:41:48] Drive: d:',
        '[2026-11-26 20:41:49] MakeMKV command: makemkvcon mkv disc:1 all "C:\Video\The Madness of King George\Disc1" --minlength=120 --progress=-same')
    $otherLog = Join-Path $tmpLogs 'other.log'
    Set-Content -LiteralPath $otherLog -Value @(
        '[2026-11-24 20:25:19] Drive: G:',
        '[2026-11-24 20:25:23] MakeMKV command: makemkvcon mkv disc:1 all "C:\Video\Michael Collins (1996)\Disc1" --minlength=120 --progress=-same')
    $idxLog = Join-Path $tmpLogs 'idx.log'
    Set-Content -LiteralPath $idxLog -Value @(
        '[2026-11-26 19:00:00] Drive Index: 0',
        '[2026-11-26 19:00:01] MakeMKV command: makemkvcon mkv disc:0 all "C:\Video\Some Film\Disc1" --minlength=120 --progress=-same')

    $win = @(
        [pscustomobject]@{ Letter = 'D:'; Name = 'GP75N'; DiscName = 'THE_MADNESS_OF_KING_GEORGE'; Busy = $false },
        [pscustomobject]@{ Letter = 'F:'; Name = 'GP75N'; DiscName = 'MEEKS_CUTOFF'; Busy = $false })
    $busy = @(Get-MakeMkvBusyDrives -Rips $rips -WinDrives $win -LogFiles @($otherLog, $kgLog))
    Assert-Equal 'D:' $busy[0].Letter 'busy drive from the rip''s own log (lower-case d: normalised), not another rip''s log'

    $busy = @(Get-MakeMkvBusyDrives -Rips $rips -WinDrives $win -LogFiles @($otherLog))
    Assert-Equal 'D:' $busy[0].Letter 'no log: the drive whose label matches the title folder'

    $sfRip = @([pscustomobject]@{ Index = 0; OutDir = 'C:\Video\Some Film\Disc1' })
    $busy = @(Get-MakeMkvBusyDrives -Rips $sfRip -WinDrives $win -LogFiles @($idxLog))
    Assert-Equal '' $busy[0].Letter 'a -DriveIndex run''s log is not trusted, and no label matches: unknown'

    $svRip = @([pscustomobject]@{ Index = 2; OutDir = 'C:\Video\Silicon Valley\Season2\Disc2' })
    $svWin = @([pscustomobject]@{ Letter = 'E:'; Name = 'x'; DiscName = 'SILICON VALLEY'; Busy = $false })
    Assert-Equal 'E:' (@(Get-MakeMkvBusyDrives -Rips $svRip -WinDrives $svWin -LogFiles @())[0].Letter) 'series folder: label matched past SeasonN\DiscN'

    $twoSame = @($win[0], [pscustomobject]@{ Letter = 'G:'; Name = 'x'; DiscName = 'The Madness of King George'; Busy = $false })
    Assert-Equal '' (@(Get-MakeMkvBusyDrives -Rips $rips -WinDrives $twoSame -LogFiles @())[0].Letter) 'two drives with matching labels: unknown, not a guess'
} finally {
    Remove-Item -LiteralPath $tmpLogs -Recurse -Force -ErrorAction SilentlyContinue
}

# Busy D: + free F: with a DVD: the existing picker switches to F:
$win[0].Busy = $true
$auto = Select-AutoVideoDrive -DrvLines $win -CurrentLetter 'D:' -GetDiscType (New-TypeMap @{ 'F:' = 'DVD-Video'; 'D:' = 'DVD-Video' })
Assert-Equal 'switch' $auto.Action 'default D: busy with MakeMKV: switch'
Assert-Equal 'F:' $auto.Drives[0].Letter 'switch to F:'
Assert-Equal $false ($script:Queried -contains 'D:') 'the busy D: is never probed'

Assert-Equal 0 (Get-FreeDriveIndex -BusyIndices @(1) -DriveCount 2) 'two drives, disc:1 busy: the free one is disc:0 (the 2026-11-26 case)'
Assert-Equal 1 (Get-FreeDriveIndex -BusyIndices @(0) -DriveCount 2) 'two drives, disc:0 busy: disc:1'
Assert-Equal 1 (Get-FreeDriveIndex -BusyIndices @(0, 2) -DriveCount 3) 'three drives, two busy: the one left'
Assert-Equal -1 (Get-FreeDriveIndex -BusyIndices @(1) -DriveCount 3) 'two free drives: not certain'
Assert-Equal -1 (Get-FreeDriveIndex -BusyIndices @(3) -DriveCount 2) 'busy index outside 0..N-1 (numbering not as expected): not certain'
Assert-Equal -1 (Get-FreeDriveIndex -BusyIndices @() -DriveCount 0) 'no drives: -1'

$ripText3 = [IO.File]::ReadAllText($ripDiscPath)
$moveOn = $ripText3.IndexOf('# ========== ANOTHER MAKEMKV RIP RUNNING: USE THE NEXT FREE DRIVE ==========')
$query = $ripText3.IndexOf('$drvProc.StartInfo.Arguments = "-r info disc:9999"')
Assert-Equal $true ($moveOn -gt 0 -and $moveOn -lt $query) 'the busy-drive move-on runs before any MakeMKV drive query'
Assert-Equal $true ($ripText3.Contains('} elseif ($script:FreeDriveIndex -ge 0) {')) 'a certain free index skips the MakeMKV query'
Assert-Equal $true ($ripText3.Contains('$drvQueryTimeoutSec = if ($preMkvRunning) { 60 } else { 20 }')) 'the query waits 60s while another rip runs'
# 2026-11-26: Enter at "[Y] Accept / [E] Edit title / [A] Abort" said "Invalid choice"
Assert-Equal $true ($ripText3.Contains('$input = Read-Host "[Y] Accept (Enter) / [E] Edit title / [A] Abort"')) 'Discovered Metadata prompt says Enter accepts'
Assert-Equal $true ($ripText3.Contains('{ $_ -in @('''', ''Y'') } { $discoveryChoice = ''Y'' }')) 'Enter (empty answer) accepts the discovered metadata'

Write-Host "`n$script:Passed passed, $script:Failed failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
