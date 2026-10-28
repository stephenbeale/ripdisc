#Requires -Version 5.1
<#
.SYNOPSIS
    Watches the optical drives and opens rip-disc.ps1 in a new window when a DVD or
    Blu-ray is inserted.

.DESCRIPTION
    Leave this running in its own window. Windows announces every newly mounted volume
    (Win32_VolumeChangeEvent, EventType 2); this script listens for that, and for an
    optical drive holding a video disc it opens a new console running
    "rip-disc.ps1 -Drive X:". Nothing else about the rip changes - rip-disc.ps1 still asks
    its usual questions in that new window.

    Nothing is polled and MakeMKV is never run from here, so the watcher cannot disturb a
    rip already in progress on another drive. The disc type is read from the volume label
    and the BDMV / VIDEO_TS folders only.

    Skipped, with a line saying why: audio CDs (use ripaudio's rip-audio.ps1), data discs,
    drives not listed in -Drive, and a drive whose previous rip-disc.ps1 run is still going.
    A disc already in a drive when the watcher starts is listed but not ripped - eject and
    re-insert it.

    Press Ctrl+C to stop watching. Rip windows already open are not affected.

.PARAMETER Drive
    Drive letters to watch (e.g. D:, E:). Default: every optical drive.

.PARAMETER RipArguments
    Extra arguments passed to rip-disc.ps1 as typed, after -Drive (e.g. "-OutputDrive F: -NoSound").

.PARAMETER SettleSeconds
    How long to wait for a newly inserted disc to become readable before giving up on it.

.PARAMETER DebounceSeconds
    Repeat "inserted" events for the same drive inside this window are ignored.

.PARAMETER RunRipFor
    Internal - what the watcher runs in each new window: rip-disc.ps1 for this drive, with
    the extra arguments read from -LockPath, which is deleted when rip-disc.ps1 returns.
    The watcher passes plain arguments like these rather than a -Command or -EncodedCommand
    string on purpose: antivirus behaviour shields treat an encoded PowerShell command as
    malware and delete the script that launched it (AVG did exactly that, 2026-10-11).

.PARAMETER LockPath
    Internal - see -RunRipFor.

.PARAMETER RipScriptPath
    Internal - the rip script to run. Default: rip-disc.ps1 next to this script.

.EXAMPLE
    .\watch-drives.ps1

.EXAMPLE
    .\watch-drives.ps1 -Drive D: -RipArguments "-OutputDrive F: -NoSound"

.EXAMPLE
    .\watch-drives.ps1 -WhatIf    # reports what would be started, opens nothing
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Drive = @(),
    [string]$RipArguments = "",
    [int]$SettleSeconds = 30,
    [int]$DebounceSeconds = 10,
    [string]$RunRipFor = "",
    [string]$LockPath = "",
    [string]$RipScriptPath = ""
)

# ========== HELPER FUNCTIONS ==========
# Same normalisation as rip-disc.ps1's function of this name (copied, per the repo's
# copy-in-each-script convention): exactly one trailing colon, whatever was typed.
function Get-NormalizedDriveLetter {
    param([string]$DriveValue)

    return ($DriveValue.Trim().TrimEnd('\')).TrimEnd(':') + ':'
}

function Write-Watch {
    param([string]$Message, [string]$Color = 'Gray')

    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

# True when $DriveLetter is one the user asked to watch. An empty list means "all".
function Test-WatchedDrive {
    param([string]$DriveLetter, [string[]]$Watch = @())

    if (-not $Watch -or $Watch.Count -eq 0) { return $true }
    return ($Watch -contains $DriveLetter)
}

# Windows can raise the arrival event more than once for a single insertion. Records the
# time in $LastSeen and returns $true when this one is a repeat inside the debounce window.
function Test-ArrivalIsDuplicate {
    param([hashtable]$LastSeen, [string]$DriveLetter, [datetime]$Now, [int]$DebounceSeconds)

    $isDuplicate = $LastSeen.ContainsKey($DriveLetter) -and
        (($Now - $LastSeen[$DriveLetter]).TotalSeconds -lt $DebounceSeconds)
    $LastSeen[$DriveLetter] = $Now
    return $isDuplicate
}

# What kind of disc sits at $Root: 'Audio CD', 'Blu-ray', 'DVD-Video', 'Data disc', or ''
# when the root cannot be read. Same rules as rip-disc.ps1's Get-DiscTypeLabel, minus its
# WMI query - the label is passed in (DriveInfo already has it) so nothing here can hang
# on another, half-disconnected drive.
function Get-DiscKind {
    param([string]$Root, [string]$VolumeLabel = "")

    if ($VolumeLabel -eq 'Audio CD') { return 'Audio CD' }
    try {
        if (-not (Test-Path -LiteralPath $Root -ErrorAction Stop)) { return '' }
    } catch {
        return ''
    }
    if (Test-Path -LiteralPath (Join-Path $Root 'BDMV') -ErrorAction SilentlyContinue) { return 'Blu-ray' }
    if (Test-Path -LiteralPath (Join-Path $Root 'VIDEO_TS') -ErrorAction SilentlyContinue) { return 'DVD-Video' }
    if (@(Get-ChildItem -LiteralPath $Root -Filter '*.cda' -File -ErrorAction SilentlyContinue).Count -gt 0) { return 'Audio CD' }
    return 'Data disc'
}

# Decides what an inserted disc of $DiscKind gets: only video discs start a rip.
function Get-InsertAction {
    param([string]$DiscKind)

    switch ($DiscKind) {
        'Blu-ray'   { return @{ Rip = $true;  Reason = '' } }
        'DVD-Video' { return @{ Rip = $true;  Reason = '' } }
        'Audio CD'  { return @{ Rip = $false; Reason = "audio CD - use ripaudio's rip-audio.ps1" } }
        'Data disc' { return @{ Rip = $false; Reason = 'data disc (no BDMV or VIDEO_TS folder)' } }
        default     { return @{ Rip = $false; Reason = 'disc could not be read' } }
    }
}

# A drive is busy while the window started for it is still open AND rip-disc.ps1 has not
# finished in it (the lock file is removed when the script returns). The window stays open
# afterwards so the summary can be read, so the process alone would say "busy" forever.
function Test-RipBusy {
    param([bool]$ProcessAlive, [bool]$LockExists)

    return ($ProcessAlive -and $LockExists)
}

# The rip-disc.ps1 call a new window makes (-RunRipFor), as text so the extra arguments
# are parsed exactly as if typed after rip-disc.ps1 at a prompt.
function Get-RipCommandText {
    param([string]$RipScriptPath, [string]$DriveLetter, [string]$RipArguments = "")

    $extra = if ([string]::IsNullOrWhiteSpace($RipArguments)) { '' } else { ' ' + $RipArguments.Trim() }
    return "& '$($RipScriptPath.Replace("'", "''"))' -Drive $DriveLetter$extra"
}

# powershell.exe arguments for the new window: this script again, in -RunRipFor mode. Only
# two paths and a drive letter go on the command line (the extra rip arguments travel in
# the lock file), so the only quoting needed is for spaces in the paths.
function Get-RipWindowArguments {
    param([string]$WatchScriptPath, [string]$DriveLetter, [string]$LockPath)

    return "-NoExit -File `"$WatchScriptPath`" -RunRipFor $DriveLetter -LockPath `"$LockPath`""
}

# Waits for the disc to spin up. Returns @{ Ready; Label } - Label is '' when not ready.
function Wait-DiscReady {
    param([string]$DriveLetter, [int]$TimeoutSec)

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        try {
            $info = [System.IO.DriveInfo]::new("$DriveLetter\")
            if ($info.IsReady) { return @{ Ready = $true; Label = [string]$info.VolumeLabel } }
        } catch { }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    return @{ Ready = $false; Label = '' }
}

function Get-OpticalDriveLetters {
    return @([System.IO.DriveInfo]::GetDrives() |
        Where-Object { $_.DriveType -eq [System.IO.DriveType]::CDRom } |
        ForEach-Object { $_.Name.TrimEnd('\') })
}

# ========== MAIN ==========
# -WhatIf is handled by hand at the one place it matters (opening a window). Left on, the
# preference leaks into the CIM module import below and prints a dozen "What if" lines.
$dryRun = [bool]$WhatIfPreference
$WhatIfPreference = $false

$ripScript = if ($RipScriptPath) { $RipScriptPath } else { Join-Path $PSScriptRoot 'rip-disc.ps1' }
if (-not (Test-Path -LiteralPath $ripScript)) {
    Write-Host "Cannot find rip-disc.ps1 ($ripScript)." -ForegroundColor Red
    exit 1
}

# ---------- -RunRipFor: the new window's side ----------
# 'exit' inside rip-disc.ps1 only ends that script when it is called with &, so the
# finally block removes the lock after a failed rip as well as a good one.
if ($RunRipFor) {
    $letter = Get-NormalizedDriveLetter $RunRipFor
    $extra = if ($LockPath -and (Test-Path -LiteralPath $LockPath)) { [string](Get-Content -LiteralPath $LockPath -Raw) } else { '' }
    try {
        & ([scriptblock]::Create((Get-RipCommandText -RipScriptPath $ripScript -DriveLetter $letter -RipArguments $extra)))
    } finally {
        if ($LockPath) { Remove-Item -LiteralPath $LockPath -ErrorAction SilentlyContinue }
    }
    return
}

# One watcher at a time: two would each open a rip window for the same disc.
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, 'Local\RipDiscWatchDrives', [ref]$createdNew)
if (-not $createdNew) {
    Write-Host "watch-drives.ps1 is already running in another window." -ForegroundColor Red
    exit 1
}

$watch = @($Drive | Where-Object { $_ } | ForEach-Object { Get-NormalizedDriveLetter $_ })
$lastSeen = @{}
$running = @{}
$sourceId = 'RipDiscDriveWatch'

try {
    $host.UI.RawUI.WindowTitle = 'RipDisc - watching drives'

    $optical = Get-OpticalDriveLetters
    $watching = @($optical | Where-Object { Test-WatchedDrive -DriveLetter $_ -Watch $watch })
    if ($watching.Count -eq 0) {
        Write-Host "No optical drive to watch (found: $(if ($optical) { $optical -join ', ' } else { 'none' }))." -ForegroundColor Red
        exit 1
    }

    Register-CimIndicationEvent -Query 'SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2' `
        -SourceIdentifier $sourceId -ErrorAction Stop | Out-Null

    Write-Watch "Watching $($watching -join ', ') - insert a DVD or Blu-ray to start a rip. Ctrl+C to stop." 'Cyan'
    foreach ($letter in $watching) {
        $loaded = Wait-DiscReady -DriveLetter $letter -TimeoutSec 0
        if ($loaded.Ready) {
            $kind = Get-DiscKind -Root "$letter\" -VolumeLabel $loaded.Label
            Write-Watch "${letter} already holds a disc ($kind) - not started. Eject and re-insert it to rip." 'Yellow'
        }
    }

    while ($true) {
        $evt = Wait-Event -SourceIdentifier $sourceId -Timeout 2
        if (-not $evt) { continue }
        Remove-Event -EventIdentifier $evt.EventIdentifier

        $letter = Get-NormalizedDriveLetter ([string]$evt.SourceEventArgs.NewEvent.DriveName)
        # The event fires for every volume (USB sticks, mounted ISOs' letters, ...).
        if ((Get-OpticalDriveLetters) -notcontains $letter) { continue }
        if (-not (Test-WatchedDrive -DriveLetter $letter -Watch $watch)) { continue }
        if (Test-ArrivalIsDuplicate -LastSeen $lastSeen -DriveLetter $letter -Now (Get-Date) -DebounceSeconds $DebounceSeconds) { continue }

        if ($running.ContainsKey($letter)) {
            $prev = $running[$letter]
            if (Test-RipBusy -ProcessAlive (-not $prev.Process.HasExited) -LockExists (Test-Path -LiteralPath $prev.LockPath)) {
                Write-Watch "${letter} disc inserted, but rip-disc.ps1 is still running for ${letter} - not started." 'Yellow'
                continue
            }
            $running.Remove($letter)
        }

        $disc = Wait-DiscReady -DriveLetter $letter -TimeoutSec $SettleSeconds
        $kind = if ($disc.Ready) { Get-DiscKind -Root "$letter\" -VolumeLabel $disc.Label } else { '' }
        $action = Get-InsertAction -DiscKind $kind
        if (-not $action.Rip) {
            Write-Watch "${letter} disc inserted - skipped: $($action.Reason)." 'Yellow'
            continue
        }

        $label = if ($disc.Label) { " '$($disc.Label)'" } else { '' }
        if ($dryRun) {
            Write-Watch "What if: ${letter} $kind$label inserted - would open a new window running: $(Get-RipCommandText -RipScriptPath $ripScript -DriveLetter $letter -RipArguments $RipArguments)" 'Cyan'
            continue
        }

        # The lock file doubles as the carrier for the extra rip arguments.
        $driveLock = Join-Path $env:TEMP "ripdisc-watch-$($letter.TrimEnd(':')).lock"
        Set-Content -LiteralPath $driveLock -Value $RipArguments
        $proc = Start-Process -FilePath 'powershell.exe' -WorkingDirectory $PSScriptRoot -PassThru `
            -ArgumentList (Get-RipWindowArguments -WatchScriptPath $PSCommandPath -DriveLetter $letter -LockPath $driveLock)
        $running[$letter] = @{ Process = $proc; LockPath = $driveLock }
        Write-Watch "${letter} $kind$label inserted - rip-disc.ps1 started in a new window." 'Green'
    }
} catch {
    Write-Host "Drive watch stopped: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue
    Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | Remove-Event
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
