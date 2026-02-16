param(
    [Parameter(Mandatory=$true)]
    [string]$title,

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [int]$Season = 0,

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$Drive = "D:",

    [Parameter()]
    [int]$DriveIndex = -1,

    [Parameter()]
    [string]$OutputDrive = "E:",

    [Parameter()]
    [switch]$Extras,

    [Parameter()]
    [switch]$Queue,

    [Parameter()]
    [switch]$Bluray,

    [Parameter()]
    [switch]$Documentary,

    [Parameter()]
    [int]$StartEpisode = 1
)

# ========== STEP TRACKING ==========
# Define the 4 processing steps
$script:AllSteps = @(
    @{ Number = 1; Name = "MakeMKV rip"; Description = "Rip disc to MKV files" }
    @{ Number = 2; Name = "HandBrake encoding"; Description = "Encode MKV to MP4" }
    @{ Number = 3; Name = "Organize files"; Description = "Rename and move files" }
    @{ Number = 4; Name = "Open directory"; Description = "Open output folder" }
)
$script:CompletedSteps = @()
$script:CurrentStep = $null
$script:LastWorkingDirectory = $null

function Set-CurrentStep {
    param([int]$StepNumber)
    $script:CurrentStep = $script:AllSteps | Where-Object { $_.Number -eq $StepNumber }
}

function Complete-CurrentStep {
    if ($script:CurrentStep) {
        $script:CompletedSteps += $script:CurrentStep
    }
}

function Get-RemainingSteps {
    $completedNumbers = $script:CompletedSteps | ForEach-Object { $_.Number }
    return $script:AllSteps | Where-Object { $_.Number -notin $completedNumbers }
}

function Get-TitleSummary {
    $contentType = if ($Documentary) { "Documentary" } elseif ($Series) { "TV Series" } else { "Movie" }
    $summary = "$contentType`: $title"
    if ($Series) {
        if ($Season -gt 0) {
            $summary += " - Season $Season, Disc $Disc"
        } else {
            $summary += " - Disc $Disc"
        }
    } elseif ($Extras) {
        $summary += " (Extras)"
    } elseif ($Disc -gt 1) {
        $summary += " (Disc $Disc - Special Features)"
    }
    return $summary
}

function Show-StepsSummary {
    param([switch]$ShowRemaining)

    Write-Host "`n--- STEPS COMPLETED ---" -ForegroundColor Green
    if ($script:CompletedSteps.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        foreach ($step in $script:CompletedSteps) {
            Write-Host "  [X] Step $($step.Number)/4: $($step.Name)" -ForegroundColor Green
        }
    }

    if ($ShowRemaining) {
        $remaining = Get-RemainingSteps
        if ($remaining.Count -gt 0) {
            Write-Host "`n--- STEPS REMAINING ---" -ForegroundColor Yellow
            foreach ($step in $remaining) {
                Write-Host "  [ ] Step $($step.Number)/4: $($step.Name) - $($step.Description)" -ForegroundColor Yellow
            }
        }
    }
}

# ========== CLOSE BUTTON PROTECTION ==========
# Disable the console window close button (X) to prevent accidental closure during rip
Add-Type -Name 'ConsoleCloseProtection' -Namespace 'Win32' -MemberDefinition @'
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
    [DllImport("user32.dll")]
    public static extern bool EnableMenuItem(IntPtr hMenu, uint uIDEnableItem, uint uEnable);
'@

$script:ConsoleWindow = [Win32.ConsoleCloseProtection]::GetConsoleWindow()
$script:ConsoleSystemMenu = [Win32.ConsoleCloseProtection]::GetSystemMenu($script:ConsoleWindow, $false)

function Disable-ConsoleClose {
    # SC_CLOSE = 0xF060, MF_BYCOMMAND = 0x0, MF_GRAYED = 0x1
    [Win32.ConsoleCloseProtection]::EnableMenuItem($script:ConsoleSystemMenu, 0xF060, 0x00000001) | Out-Null
}

function Enable-ConsoleClose {
    # SC_CLOSE = 0xF060, MF_BYCOMMAND = 0x0, MF_ENABLED = 0x0
    [Win32.ConsoleCloseProtection]::EnableMenuItem($script:ConsoleSystemMenu, 0xF060, 0x00000000) | Out-Null
}

# ========== HELPER FUNCTIONS ==========
function Get-UniqueFilePath {
    param([string]$DestDir, [string]$FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $targetPath = Join-Path $DestDir $FileName

    if (!(Test-Path $targetPath)) {
        return $targetPath
    }

    $counter = 1
    do {
        $newName = "$baseName-$counter$extension"
        $targetPath = Join-Path $DestDir $newName
        $counter++
    } while (Test-Path $targetPath)

    return $targetPath
}

function Test-DriveReady {
    param([string]$Path)

    # Extract the drive letter from the path (e.g., "E:" from "E:\DVDs\Movie")
    $driveLetter = [System.IO.Path]::GetPathRoot($Path)
    if (-not $driveLetter) {
        return @{ Ready = $false; Drive = "Unknown"; Message = "Could not determine drive letter from path: $Path" }
    }

    # Normalize drive letter (remove trailing backslash for display)
    $driveDisplay = $driveLetter.TrimEnd('\')

    # Check if the drive exists and is ready
    try {
        $drive = Get-PSDrive -Name $driveDisplay.TrimEnd(':') -ErrorAction Stop
        if ($drive) {
            # Additional check: try to access the drive root
            if (Test-Path $driveLetter -ErrorAction SilentlyContinue) {
                return @{ Ready = $true; Drive = $driveDisplay; Message = "Drive is ready" }
            } else {
                return @{ Ready = $false; Drive = $driveDisplay; Message = "Destination drive $driveDisplay is not ready - please ensure the drive is connected and mounted" }
            }
        }
    } catch {
        return @{ Ready = $false; Drive = $driveDisplay; Message = "Destination drive $driveDisplay is not ready - please ensure the drive is connected and mounted" }
    }

    return @{ Ready = $false; Drive = $driveDisplay; Message = "Destination drive $driveDisplay is not ready - please ensure the drive is connected and mounted" }
}

function Write-Log {
    param([string]$Message)
    if ($script:LogFile) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "[$timestamp] $Message"
        Add-Content -Path $script:LogFile -Value $entry
    }
}

# ========== DRIVE CONFIRMATION ==========
# Show which drive will be used and confirm before proceeding
$driveLetter = if ($Drive -match ':$') { $Drive } else { "${Drive}:" }
$driveDescription = if ($DriveIndex -ge 0) {
    $hint = switch ($DriveIndex) {
        0 { "D: internal" }
        1 { "G: ASUS external" }
        default { "unknown drive" }
    }
    "Drive Index $DriveIndex ($hint)"
} else {
    "Drive $driveLetter"
}
# ========== TITLE VALIDATION ==========
# Warn if title appears to contain metadata that should be separate parameters
$titleWarnings = @()
if ($Series) {
    if ($title -match '(?i)\bseries\s*\d') {
        $titleWarnings += "Contains 'Series N' - use -Season parameter instead"
    }
    if ($title -match '(?i)\bseason\s*\d') {
        $titleWarnings += "Contains 'Season N' - use -Season parameter instead"
    }
    if ($title -match '(?i)\bdisc\s*\d') {
        $titleWarnings += "Contains 'Disc N' - use -Disc parameter instead"
    }
    if ($title -match '(?i)\bS\d{1,2}E\d') {
        $titleWarnings += "Contains episode code (e.g. S01E01) - use -Series -Season instead"
    }
}
if ($titleWarnings.Count -gt 0) {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "WARNING: Title may contain misplaced metadata" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Title: `"$title`"" -ForegroundColor Yellow
    foreach ($w in $titleWarnings) {
        Write-Host "  ! $w" -ForegroundColor Yellow
    }
    Write-Host "`nExpected usage:" -ForegroundColor Cyan
    Write-Host "  .\rip-disc.ps1 -title `"Fargo`" -Series -Season 1 -Disc 2" -ForegroundColor White
    Write-Host ""
    $continueChoice = Read-Host "Continue with this title? (y/N)"
    if ($continueChoice -ne 'y' -and $continueChoice -ne 'Y') {
        Write-Host "Aborted. Please re-run with correct parameters." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Ready to rip: $title" -ForegroundColor White
if ($Documentary) {
    $discType = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: Documentary - $discType$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
} elseif ($Series) {
    if ($Season -gt 0) {
        $seasonTagPreview = "S{0:D2}" -f $Season
        Write-Host "Type: TV Series - Season $Season ($seasonTagPreview), Disc $Disc" -ForegroundColor White
    } else {
        Write-Host "Type: TV Series - Disc $Disc (no season folder)" -ForegroundColor White
    }
} else {
    $discType = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: Movie - $discType$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
}
Write-Host "Using: $driveDescription" -ForegroundColor Yellow
Write-Host "Output Drive: $OutputDrive" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
$host.UI.RawUI.WindowTitle = "rip-disc - INPUT"
$response = Read-Host "Press Enter to continue, or Ctrl+C to abort"

# Disable close button to prevent accidental window closure during rip
Disable-ConsoleClose

# ========== SET WINDOW TITLE ==========
# Set PowerShell window title to help identify concurrent rips
# Title comes FIRST so it's visible in narrow terminal tabs
if ($Series) {
    $windowTitle = "$title"
    if ($Season -gt 0) { $windowTitle += " S$Season" }
    $windowTitle += " Disc $Disc"
} else {
    $windowTitle = "$title"
    if ($Extras -or $Disc -gt 1) { $windowTitle += "-extras" }
}
$host.UI.RawUI.WindowTitle = $windowTitle

# ========== CONFIGURATION ==========
# MakeMKV temp directory - use subdirectory for multi-disc and extras rips
if ($Extras) {
    $makemkvOutputDir = "C:\Video\$title\Extras"
} elseif ($Series -and $Season -gt 0) {
    $makemkvOutputDir = "C:\Video\$title\Season$Season\Disc$Disc"
} else {
    $makemkvOutputDir = "C:\Video\$title\Disc$Disc"
}

# Normalize output drive letter (add colon if missing)
$outputDriveLetter = if ($OutputDrive -match ':$') { $OutputDrive } else { "${OutputDrive}:" }

# Documentaries: organize into Documentaries folder
# Series: organize into Season subfolders (only if Season explicitly specified)
# Movies: organize into title folder with optional extras
if ($Documentary) {
    $finalOutputDir = "$outputDriveLetter\Documentaries\$title"
} elseif ($Series) {
    $seriesBaseDir = "$outputDriveLetter\Series\$title"
    if ($Season -gt 0) {
        # Season explicitly specified - use Season subfolder
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $finalOutputDir = Join-Path $seriesBaseDir $seasonFolder
    } else {
        # No season specified - output directly to series folder, no season tag
        $seasonTag = $null
        $finalOutputDir = $seriesBaseDir
    }
} else {
    $finalOutputDir = "$outputDriveLetter\DVDs\$title"
}

$makemkvconPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"
$handbrakePath = "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe"

# ========== LOGGING SETUP ==========
$logDir = "C:\Video\logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDiscLabel = if ($Extras) { "extras" } else { "disc${Disc}" }
$script:LogFile = Join-Path $logDir "${title}_${logDiscLabel}_${logTimestamp}.log"

Write-Log "========== RIP SESSION STARTED =========="
Write-Log "Title: $title"
Write-Log "Type: $(if ($Documentary) { 'Documentary' } elseif ($Series) { 'TV Series' } else { 'Movie' })"
Write-Log "Disc: $Disc$(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1 -and -not $Series) { ' (Special Features)' })"
if ($Series -and $Season -gt 0) {
    Write-Log "Season: $Season"
}
if ($DriveIndex -ge 0) {
    Write-Log "Drive Index: $DriveIndex"
} else {
    Write-Log "Drive: $driveLetter"
}
Write-Log "Output Drive: $outputDriveLetter"
Write-Log "MakeMKV Output: $makemkvOutputDir"
Write-Log "Final Output: $finalOutputDir"
Write-Log "Log file: $($script:LogFile)"

function Stop-WithError {
    param([string]$Step, [string]$Message)

    $host.UI.RawUI.WindowTitle = "$($host.UI.RawUI.WindowTitle) - ERROR"

    # Log the error
    Write-Log "========== ERROR =========="
    Write-Log "Failed at: $Step"
    Write-Log "Message: $Message"
    if ($script:CompletedSteps.Count -gt 0) {
        Write-Log "Completed steps: $(($script:CompletedSteps | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    } else {
        Write-Log "Completed steps: (none)"
    }
    $remaining = Get-RemainingSteps
    if ($remaining.Count -gt 0) {
        Write-Log "Remaining steps: $(($remaining | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    }
    Write-Log "Log file: $($script:LogFile)"

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red

    # Always show what was being processed
    Write-Host "`nProcessing: $(Get-TitleSummary)" -ForegroundColor White

    Write-Host "`nError at: $Step" -ForegroundColor Red
    Write-Host "Message: $Message" -ForegroundColor Red

    # Show completed and remaining steps
    Show-StepsSummary -ShowRemaining

    # Determine which directory to open (where leftover files might be)
    $directoryToOpen = $null
    if ($script:LastWorkingDirectory -and (Test-Path $script:LastWorkingDirectory)) {
        $directoryToOpen = $script:LastWorkingDirectory
    } elseif (Test-Path $makemkvOutputDir) {
        $directoryToOpen = $makemkvOutputDir
    } elseif (Test-Path $finalOutputDir) {
        $directoryToOpen = $finalOutputDir
    }

    # Show manual steps the user needs to handle
    Write-Host "`n--- MANUAL STEPS NEEDED ---" -ForegroundColor Cyan
    $remaining = Get-RemainingSteps
    foreach ($step in $remaining) {
        switch ($step.Number) {
            1 { Write-Host "  - Re-run MakeMKV to rip the disc" -ForegroundColor Yellow }
            2 {
                Write-Host "  - Encode MKV files with HandBrake" -ForegroundColor Yellow
                if (Test-Path $makemkvOutputDir) {
                    Write-Host "    MKV files location: $makemkvOutputDir" -ForegroundColor Gray
                }
            }
            3 {
                Write-Host "  - Rename files to proper format" -ForegroundColor Yellow
                if ($Series) {
                    Write-Host "    Format: $title-originalname.mp4" -ForegroundColor Gray
                } else {
                    if ($isMainFeatureDisc) {
                        Write-Host "    Format: $title-Feature.mp4 (largest file)" -ForegroundColor Gray
                        Write-Host "    Move extras to: $extrasDir" -ForegroundColor Gray
                    } else {
                        Write-Host "    Format: $title-Special Features-originalname.mp4" -ForegroundColor Gray
                        Write-Host "    Move all files to: $extrasDir" -ForegroundColor Gray
                    }
                }
            }
            4 { Write-Host "  - Open output directory to verify files" -ForegroundColor Yellow }
        }
    }

    # Open the relevant directory so user can see leftover files
    if ($directoryToOpen) {
        Write-Host "`n--- OPENING DIRECTORY ---" -ForegroundColor Cyan
        Write-Host "Opening: $directoryToOpen" -ForegroundColor Yellow
        Write-Host "(This is where leftover/partial files may be located)" -ForegroundColor Gray
        Start-Process explorer.exe -ArgumentList $directoryToOpen
    }

    # Show recovery script path if it exists (encoding failed mid-way)
    if ($recoveryScriptPath -and (Test-Path $recoveryScriptPath)) {
        Write-Host "`n--- RECOVERY SCRIPT ---" -ForegroundColor Cyan
        Write-Host "Recovery script: $recoveryScriptPath" -ForegroundColor Yellow
        Write-Host "Run this to encode remaining files: .\$(Split-Path $recoveryScriptPath -Leaf)" -ForegroundColor White
        Write-Log "Recovery script available: $recoveryScriptPath"
    }

    Write-Host "`nLog file: $($script:LogFile)" -ForegroundColor Yellow
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "Please complete the remaining steps manually" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    Enable-ConsoleClose
    exit 1
}

$contentType = if ($Documentary) { "Documentary" } elseif ($Series) { "TV Series" } else { "Movie" }
# Documentaries are treated like movies for file organization (Feature file, extras subfolder)
$isMainFeatureDisc = (-not $Series) -and ($Disc -eq 1) -and (-not $Extras)
$extrasDir = Join-Path $finalOutputDir "extras"

# For disc 2+, ensure parent dir and extras folder exist upfront (disc 1 may still be running)
if (-not $isMainFeatureDisc -and -not $Series) {
    # Check if destination drive is ready before attempting to create directories
    $driveCheck = Test-DriveReady -Path $finalOutputDir
    if (-not $driveCheck.Ready) {
        Write-Host "`nERROR: $($driveCheck.Message)" -ForegroundColor Red
        exit 1
    }
    try {
        if (!(Test-Path $finalOutputDir)) {
            New-Item -ItemType Directory -Path $finalOutputDir -Force -ErrorAction Stop | Out-Null
        }
        if (!(Test-Path $extrasDir)) {
            New-Item -ItemType Directory -Path $extrasDir -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Host "`nERROR: Cannot create output directory - $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DVD/Blu-ray Ripping & Encoding Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Title: $title" -ForegroundColor White
Write-Host "Type: $contentType" -ForegroundColor White
if ($DriveIndex -ge 0) {
    $driveHint = switch ($DriveIndex) {
        0 { "D: internal" }
        1 { "G: ASUS external" }
        default { "unknown" }
    }
    Write-Host "Drive Index: $DriveIndex ($driveHint)" -ForegroundColor White
} else {
    Write-Host "Drive: $driveLetter" -ForegroundColor White
}
Write-Host "Output Drive: $outputDriveLetter" -ForegroundColor White
if ($Series) {
    if ($Season -gt 0) {
        Write-Host "Season: $Season ($seasonTag)" -ForegroundColor White
    } else {
        Write-Host "Season: (none - no season folder)" -ForegroundColor White
    }
    Write-Host "Disc: $Disc" -ForegroundColor White
} else {
    Write-Host "Disc: $Disc$(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1) { ' (Special Features)' })" -ForegroundColor White
}
Write-Host "MakeMKV Output: $makemkvOutputDir" -ForegroundColor White
Write-Host "Final Output: $finalOutputDir" -ForegroundColor White
Write-Host "Log file: $($script:LogFile)" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

# ========== STEP 1: RIP WITH MAKEMKV ==========
Set-CurrentStep -StepNumber 1
$script:LastWorkingDirectory = $makemkvOutputDir
Write-Log "STEP 1/4: Starting MakeMKV rip..."
Write-Host "[STEP 1/4] Starting MakeMKV rip..." -ForegroundColor Green

# Use disc: syntax with index if provided (completely bypasses drive enumeration)
# Otherwise fall back to dev: syntax which still enumerates drives
if ($DriveIndex -ge 0) {
    $discSource = "disc:$DriveIndex"
    Write-Host "Using drive index: $DriveIndex (bypasses drive enumeration)" -ForegroundColor Green
} else {
    $discSource = "dev:$driveLetter"
    Write-Host "Using drive: $driveLetter (may enumerate other drives)" -ForegroundColor Yellow
    Write-Host "Tip: Use -DriveIndex to bypass drive enumeration" -ForegroundColor Gray
}

Write-Host "Creating directory: $makemkvOutputDir" -ForegroundColor Yellow
if (Test-Path $makemkvOutputDir) {
    $existingFiles = Get-ChildItem -Path $makemkvOutputDir -File -ErrorAction SilentlyContinue
    if ($existingFiles -and $existingFiles.Count -gt 0) {
        Write-Host "`nWARNING: Directory already exists with $($existingFiles.Count) file(s):" -ForegroundColor Yellow
        Write-Host "  $makemkvOutputDir" -ForegroundColor White
        foreach ($ef in $existingFiles) {
            Write-Host "  - $($ef.Name) ($([math]::Round($ef.Length/1GB, 2)) GB)" -ForegroundColor Gray
        }

        # Find the next available suffix
        $suffix = 1
        while (Test-Path "${makemkvOutputDir}-${suffix}") { $suffix++ }
        $suffixedDir = "${makemkvOutputDir}-${suffix}"

        Write-Host "`nChoose an option:" -ForegroundColor Cyan
        Write-Host "  [1] Delete existing files and reuse directory" -ForegroundColor Yellow
        Write-Host "  [2] Use suffixed directory: $suffixedDir" -ForegroundColor Yellow

        $choice = $null
        while ($choice -ne '1' -and $choice -ne '2') {
            $choice = Read-Host "Enter 1 or 2"
            if ($choice -ne '1' -and $choice -ne '2') {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        }

        if ($choice -eq '1') {
            Write-Host "Deleting existing files..." -ForegroundColor Yellow
            $existingFiles | Remove-Item -Force
            Write-Host "Deleted $($existingFiles.Count) existing file(s)" -ForegroundColor Green
            Write-Log "User chose to delete $($existingFiles.Count) existing file(s) in $makemkvOutputDir"
        } else {
            $makemkvOutputDir = $suffixedDir
            New-Item -ItemType Directory -Path $makemkvOutputDir -Force | Out-Null
            Write-Host "Using suffixed directory: $makemkvOutputDir" -ForegroundColor Green
            Write-Log "User chose suffixed directory: $makemkvOutputDir"
        }
    } else {
        Write-Host "Directory exists (empty)" -ForegroundColor Gray
    }
} else {
    New-Item -ItemType Directory -Path $makemkvOutputDir | Out-Null
    Write-Host "Directory created successfully" -ForegroundColor Green
}

Write-Host "`nExecuting MakeMKV command..." -ForegroundColor Yellow
Write-Host "Command: makemkvcon mkv $discSource all `"$makemkvOutputDir`" --minlength=120" -ForegroundColor Gray
Write-Log "MakeMKV command: makemkvcon mkv $discSource all `"$makemkvOutputDir`" --minlength=120"

# Capture MakeMKV output to analyze for specific errors
$makemkvOutput = & $makemkvconPath mkv $discSource all $makemkvOutputDir --minlength=120 2>&1 | Tee-Object -Variable makemkvFullOutput
$makemkvExitCode = $LASTEXITCODE
$makemkvOutputText = $makemkvFullOutput -join "`n"

# Check if MakeMKV succeeded - provide specific error messages for common issues
if ($makemkvExitCode -ne 0) {
    # Analyze output to determine the specific error
    $errorMessage = "MakeMKV exited with code $makemkvExitCode"

    # Check for drive not found / doesn't exist
    if ($makemkvOutputText -match "Failed to open disc" -or
        $makemkvOutputText -match "no disc" -or
        $makemkvOutputText -match "can't find" -or
        $makemkvOutputText -match "invalid drive") {
        if ($DriveIndex -ge 0) {
            $errorMessage = "Drive not found: Drive index $DriveIndex does not exist or is not accessible"
        } else {
            $errorMessage = "Drive not found: $driveLetter - verify the drive letter is correct"
        }
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    }
    # Check for empty drive / no disc inserted
    elseif ($makemkvOutputText -match "no media" -or
            $makemkvOutputText -match "medium not present" -or
            $makemkvOutputText -match "drive is empty" -or
            $makemkvOutputText -match "no disc in drive" -or
            $makemkvOutputText -match "insert a disc") {
        if ($DriveIndex -ge 0) {
            $driveHintMsg = switch ($DriveIndex) {
                0 { "D: internal" }
                1 { "G: ASUS external" }
                default { "drive index $DriveIndex" }
            }
            $errorMessage = "Drive is empty ($driveHintMsg) - please insert a disc"
        } else {
            $errorMessage = "Drive $driveLetter is empty - please insert a disc"
        }
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    }
    # Check for disc not readable / can't detect disc
    elseif ($makemkvOutputText -match "can't access" -or
            $makemkvOutputText -match "read error" -or
            $makemkvOutputText -match "cannot read" -or
            $makemkvOutputText -match "failed to read") {
        $errorMessage = "No disc detected in drive - the disc may be damaged or unreadable"
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    }

    Stop-WithError -Step "STEP 1/4: MakeMKV rip" -Message $errorMessage
}

$rippedFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
if ($null -eq $rippedFiles -or $rippedFiles.Count -eq 0) {
    # MakeMKV succeeded but no files created - likely no valid titles found
    $errorMessage = "No MKV files were created"

    # Check output for clues about why no files were created
    if ($makemkvOutputText -match "no valid" -or $makemkvOutputText -match "0 titles") {
        $errorMessage = "No disc detected in drive - MakeMKV could not find any valid titles"
    } elseif ($makemkvOutputText -match "copy protection" -or $makemkvOutputText -match "protected") {
        $errorMessage = "Disc may be copy-protected or encrypted - MakeMKV could not extract titles"
    } else {
        $errorMessage = "No MKV files were created - check if disc is readable and contains valid content"
    }

    Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    Stop-WithError -Step "STEP 1/4: MakeMKV rip" -Message $errorMessage
}

Write-Host "`nMakeMKV rip complete!" -ForegroundColor Green
Write-Host "Files ripped: $($rippedFiles.Count)" -ForegroundColor White
Write-Log "STEP 1/4: MakeMKV rip complete - $($rippedFiles.Count) file(s)"
foreach ($file in $rippedFiles) {
    Write-Host "  - $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)" -ForegroundColor Gray
    Write-Log "  Ripped: $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)"
}
Complete-CurrentStep

# Eject disc (with timeout to prevent hanging if drive is busy)
Write-Host "`nEjecting disc from drive $driveLetter..." -ForegroundColor Yellow
$ejectJob = Start-Job -ScriptBlock {
    param($drive)
    $shell = New-Object -comObject Shell.Application
    $shell.Namespace(17).ParseName($drive).InvokeVerb("Eject")
} -ArgumentList $driveLetter
$ejectCompleted = $ejectJob | Wait-Job -Timeout 15
if ($ejectCompleted) {
    Remove-Job $ejectJob -Force
    Write-Host "Disc ejected successfully" -ForegroundColor Green
    Write-Log "Disc ejected from drive $driveLetter"
} else {
    Stop-Job $ejectJob
    Remove-Job $ejectJob -Force
    Write-Host "Disc eject timed out - please eject manually" -ForegroundColor Yellow
    Write-Log "WARNING: Disc eject timed out for drive $driveLetter"
}


# ========== QUEUE MODE: ADD TO QUEUE AND EXIT ==========
if ($Queue) {
    Write-Log "QUEUE MODE: Writing encoding job to queue file..."
    Write-Host "`n[QUEUE MODE] Adding encoding job to queue..." -ForegroundColor Green

    $queueFilePath = "C:\Video\handbrake-queue.json"
    $lockFilePath = "$queueFilePath.lock"

    $entry = @{
        Title = $title
        Series = [bool]$Series
        Season = $Season
        Disc = $Disc
        OutputDrive = $OutputDrive
        QueuedAt = (Get-Date -Format "o")
    }

    # Read existing queue with file locking
    $retryCount = 0
    $maxRetries = 10
    $lockAcquired = $false

    while (-not $lockAcquired -and $retryCount -lt $maxRetries) {
        try {
            $lockStream = [System.IO.File]::Open($lockFilePath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $lockAcquired = $true
        } catch {
            $retryCount++
            Start-Sleep -Milliseconds 500
        }
    }

    if (-not $lockAcquired) {
        Write-Host "WARNING: Could not acquire lock file - writing without lock" -ForegroundColor Red
    }

    try {
        if (Test-Path $queueFilePath) {
            $queue = Get-Content $queueFilePath -Raw | ConvertFrom-Json
            if ($queue -isnot [System.Array]) { $queue = @($queue) }
        } else {
            $queue = @()
        }

        $queue += $entry
        $queue | ConvertTo-Json -Depth 10 | Set-Content $queueFilePath -Encoding UTF8
    } finally {
        if ($lockStream) { $lockStream.Close() }
        Remove-Item $lockFilePath -Force -ErrorAction SilentlyContinue
    }

    $mkvCount = (Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv").Count

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "QUEUED!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nTitle: $title" -ForegroundColor White
    Write-Host "MKV files: $mkvCount" -ForegroundColor White
    Write-Host "Queue file: $queueFilePath" -ForegroundColor White
    Write-Host "Total jobs in queue: $($queue.Count)" -ForegroundColor White
    Write-Host "`nRun 'RipDisc -processQueue' to encode all queued jobs sequentially" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Cyan

    Write-Log "QUEUE MODE: Job added to queue ($($queue.Count) total jobs)"
    Write-Log "Queue file: $queueFilePath"

    Enable-ConsoleClose
    $host.UI.RawUI.WindowTitle = "$windowTitle - QUEUED"
    exit 0
}


# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
Set-CurrentStep -StepNumber 2
$script:LastWorkingDirectory = $finalOutputDir
Write-Log "STEP 2/4: Starting HandBrake encoding..."
Write-Host "`n[STEP 2/4] Starting HandBrake encoding..." -ForegroundColor Green

# Check if destination drive is ready before attempting to create directories
Write-Host "Checking destination drive..." -ForegroundColor Yellow
$driveCheck = Test-DriveReady -Path $finalOutputDir
if (-not $driveCheck.Ready) {
    Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message $driveCheck.Message
}
Write-Host "Destination drive $($driveCheck.Drive) is ready" -ForegroundColor Green

Write-Host "Creating directory: $finalOutputDir" -ForegroundColor Yellow
if (!(Test-Path $finalOutputDir)) {
    try {
        New-Item -ItemType Directory -Path $finalOutputDir -ErrorAction Stop | Out-Null
        Write-Host "Directory created successfully" -ForegroundColor Green
    } catch {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "Cannot create output directory: $finalOutputDir - $($_.Exception.Message)"
    }
} else {
    Write-Host "Directory already exists" -ForegroundColor Yellow
}

$mkvFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv"

# Series mode: detect and skip composite mega-file (all episodes in one)
if ($Series -and $mkvFiles.Count -ge 3) {
    $sortedBySize = $mkvFiles | Sort-Object Length -Descending
    $largest = $sortedBySize[0]
    $secondLargest = $sortedBySize[1]
    if ($largest.Length -ge ($secondLargest.Length * 2)) {
        Write-Host "`nComposite file detected (skipping encode): $($largest.Name) ($([math]::Round($largest.Length/1GB, 2)) GB)" -ForegroundColor Yellow
        Write-Log "Skipping composite file: $($largest.Name) ($([math]::Round($largest.Length/1GB, 2)) GB)"
        $mkvFiles = $mkvFiles | Where-Object { $_.FullName -ne $largest.FullName }
        Write-Host "Encoding $($mkvFiles.Count) episode file(s)" -ForegroundColor Green
    }
}

# ========== GENERATE RECOVERY SCRIPT ==========
# Create a recovery .ps1 with HandBrakeCLI commands for each MKV file.
# If encoding fails, the user can run this script to resume encoding manually.
$safeTitle = $title -replace '[\\/:*?"<>|]', '_'
$dateStamp = Get-Date -Format "yyyy-MM-dd"
$recoveryScriptPath = "C:\Video\recovery_${safeTitle}_${dateStamp}.ps1"
$recoveryLines = @(
    "# HandBrake recovery script for: $title"
    "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "# Run this script to encode any remaining MKV files that were not successfully encoded."
    "# Already-encoded files (existing .mp4) will be skipped automatically."
    ""
    "`$handbrakePath = `"$handbrakePath`""
    "`$finalOutputDir = `"$finalOutputDir`""
    ""
    "if (!(Test-Path `$finalOutputDir)) { New-Item -ItemType Directory -Path `$finalOutputDir -Force | Out-Null }"
    ""
)
foreach ($mkv in $mkvFiles) {
    $inputFile = $mkv.FullName
    $outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")
    $recoveryLines += "# --- $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB) ---"
    $recoveryLines += "if (!(Test-Path `"$outputFile`")) {"
    if ($Bluray) {
        $recoveryLines += "    Write-Host `"Encoding: $($mkv.Name)`" -ForegroundColor Cyan"
        $recoveryLines += "    & `$handbrakePath -i `"$inputFile`" -o `"$outputFile`" --preset `"Fast 1080p30`" --all-audio --all-subtitles --subtitle-burned=none --verbose=1"
        $recoveryLines += "    if (`$LASTEXITCODE -ne 0) {"
        $recoveryLines += "        Write-Host `"Subtitle encoding failed - retrying without subtitles...`" -ForegroundColor Yellow"
        $recoveryLines += "        if (Test-Path `"$outputFile`") { Remove-Item `"$outputFile`" -Force }"
        $recoveryLines += "        & `$handbrakePath -i `"$inputFile`" -o `"$outputFile`" --preset `"Fast 1080p30`" --all-audio --verbose=1"
        $recoveryLines += "    }"
    } else {
        $recoveryLines += "    Write-Host `"Encoding: $($mkv.Name)`" -ForegroundColor Cyan"
        $recoveryLines += "    & `$handbrakePath -i `"$inputFile`" -o `"$outputFile`" --preset `"Fast 1080p30`" --all-audio --all-subtitles --subtitle-burned=none --verbose=1"
    }
    $recoveryLines += "} else {"
    $recoveryLines += "    Write-Host `"Skipping (already encoded): $($mkv.Name)`" -ForegroundColor Gray"
    $recoveryLines += "}"
    $recoveryLines += ""
}
$recoveryLines += "Write-Host `"`nRecovery encoding complete.`" -ForegroundColor Green"
$recoveryLines | Set-Content -Path $recoveryScriptPath -Encoding UTF8
Write-Host "Recovery script: $recoveryScriptPath" -ForegroundColor Yellow
Write-Log "Recovery script created: $recoveryScriptPath"

$fileCount = 0
foreach ($mkv in $mkvFiles) {
    $fileCount++
    $inputFile = $mkv.FullName
    $outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")

    Write-Host "`n--- Encoding file $fileCount of $($mkvFiles.Count) ---" -ForegroundColor Cyan
    Write-Host "Input:  $($mkv.Name)" -ForegroundColor White
    Write-Host "Output: $($mkv.BaseName).mp4" -ForegroundColor White
    Write-Host "Size:   $([math]::Round($mkv.Length/1GB, 2)) GB" -ForegroundColor White
    Write-Log "Encoding file $fileCount of $($mkvFiles.Count): $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB)"

    Write-Host "`nExecuting HandBrake..." -ForegroundColor Yellow

    # Always try with subtitles first (not burned in)
    $handbrakeArgs = @(
        "-i", $inputFile,
        "-o", $outputFile,
        "--preset", "Fast 1080p30",
        "--all-audio",
        "--all-subtitles",
        "--subtitle-burned=none",
        "--verbose=1"
    )
    & $handbrakePath @handbrakeArgs
    $handbrakeExitCode = $LASTEXITCODE

    # For Bluray: if subtitle encoding fails, retry without subtitles (PGS incompatibility)
    if ($handbrakeExitCode -ne 0 -and $Bluray) {
        Write-Host "`nBluray subtitle encoding failed - retrying without subtitles..." -ForegroundColor Yellow
        Write-Log "Bluray subtitle encoding failed for $($mkv.Name) - retrying without subtitles"

        # Delete partial output if exists
        if (Test-Path $outputFile) {
            Remove-Item $outputFile -Force
        }

        $handbrakeArgsNoSubs = @(
            "-i", $inputFile,
            "-o", $outputFile,
            "--preset", "Fast 1080p30",
            "--all-audio",
            "--verbose=1"
        )
        & $handbrakePath @handbrakeArgsNoSubs
        $handbrakeExitCode = $LASTEXITCODE
    }

    if ($handbrakeExitCode -ne 0) {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "HandBrake exited with code $handbrakeExitCode while encoding $($mkv.Name)"
    }

    if (Test-Path $outputFile) {
        $encodedSize = (Get-Item $outputFile).Length
        Write-Host "`nEncoding complete: $($mkv.Name)" -ForegroundColor Green
        Write-Host "Output size: $([math]::Round($encodedSize/1GB, 2)) GB" -ForegroundColor White
        Write-Log "Encoded: $($mkv.Name) -> $($mkv.BaseName).mp4 ($([math]::Round($encodedSize/1GB, 2)) GB)"
    } else {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "Output file not created for $($mkv.Name)"
    }
}
Complete-CurrentStep
Write-Log "STEP 2/4: HandBrake encoding complete - $fileCount file(s) encoded"

# Delete recovery script after successful encoding
if (Test-Path $recoveryScriptPath) {
    Remove-Item $recoveryScriptPath -Force
    Write-Host "Recovery script deleted (encoding successful)" -ForegroundColor Gray
    Write-Log "Recovery script deleted: $recoveryScriptPath"
}

# Wait for HandBrake to fully release file handles before proceeding
Write-Host "`nWaiting for file handles to be released..." -ForegroundColor Yellow
Start-Sleep -Seconds 3
Write-Host "File handle wait complete" -ForegroundColor Green

# Delete MakeMKV temporary directory after successful encode
Write-Host "`nChecking for successful encodes..." -ForegroundColor Yellow
$encodedFiles = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4"
$script:EncodedFilesTooSmall = $false
if ($encodedFiles.Count -gt 0) {
    # Safety check: verify largest encoded file is at least 100MB
    # If all files are suspiciously small, encoding likely failed silently
    $largestEncoded = $encodedFiles | Sort-Object Length -Descending | Select-Object -First 1
    $largestSizeMB = [math]::Round($largestEncoded.Length / 1MB, 2)
    $minSizeMB = 100

    if ($largestSizeMB -lt $minSizeMB) {
        $script:EncodedFilesTooSmall = $true
        Write-Host "WARNING: Largest encoded file is only $largestSizeMB MB (threshold: $minSizeMB MB)" -ForegroundColor Red
        Write-Host "Encoded files may be corrupt - keeping MakeMKV source files for safety" -ForegroundColor Red
        Write-Host "Source MKV directory preserved: $makemkvOutputDir" -ForegroundColor Yellow
        Write-Log "WARNING: Largest encoded file ($($largestEncoded.Name)) is only $largestSizeMB MB - below $minSizeMB MB safety threshold"
        Write-Log "Keeping MakeMKV source directory: $makemkvOutputDir"
        # Open Recycle Bin so user can review
        Start-Process explorer.exe -ArgumentList "shell:RecycleBinFolder"
        Write-Host "Opened Recycle Bin for review" -ForegroundColor Yellow
        Write-Log "Opened Recycle Bin for user review"
    } else {
        Write-Host "Found $($encodedFiles.Count) encoded file(s) (largest: $largestSizeMB MB)" -ForegroundColor Green
        Write-Host "Removing temporary MakeMKV directory: $makemkvOutputDir" -ForegroundColor Yellow
        Remove-Item -Path $makemkvOutputDir -Recurse -Force
        Write-Host "Temporary files removed successfully" -ForegroundColor Green
        Write-Log "Temporary MKV directory removed: $makemkvOutputDir"
    }
} else {
    Write-Host "WARNING: No encoded files found. Keeping MakeMKV directory." -ForegroundColor Red
    Write-Log "WARNING: No encoded files found - keeping MakeMKV directory"
}

# ========== STEP 3: RENAME AND ORGANIZE ==========
Set-CurrentStep -StepNumber 3
$script:LastWorkingDirectory = $finalOutputDir
Write-Log "STEP 3/4: Organizing files..."
Write-Host "`n[STEP 3/4] Organizing files..." -ForegroundColor Green
cd $finalOutputDir

Write-Host "Current directory: $finalOutputDir" -ForegroundColor Yellow

# delete image files first (only if they exist)
Write-Host "`nDeleting image files..." -ForegroundColor Yellow
$imageFiles = Get-ChildItem -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png|gif|bmp)$' }
if ($imageFiles) {
    Write-Host "Image files to delete: $($imageFiles.Count)" -ForegroundColor White
    $imageFiles | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
    $imageFiles | Remove-Item -ErrorAction SilentlyContinue
    Write-Host "Image files deleted" -ForegroundColor Green
} else {
    Write-Host "No image files found" -ForegroundColor Gray
}

if ($Series) {
    # ========== SERIES MODE: Rename to Jellyfin episode format ==========
    Write-Host "`nRenaming episodes to Jellyfin format..." -ForegroundColor Yellow
    $seasonTag = if ($Season -gt 0) { "S{0:D2}" -f $Season } else { "" }

    # Get video files sorted by name (MakeMKV title order = episode order)
    $episodeFiles = Get-ChildItem -File | Where-Object {
        $_.Extension -match '\.(mp4|mkv)$'
    } | Sort-Object Name

    $episodeNum = $StartEpisode
    foreach ($file in $episodeFiles) {
        $episodeTag = "E{0:D2}" -f $episodeNum
        $newName = "$title-$seasonTag$episodeTag$($file.Extension)"
        Write-Host "  $($file.Name) -> $newName" -ForegroundColor Gray
        $maxRetries = 5
        $retryDelay = 3
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
                break
            } catch [System.IO.IOException] {
                if ($attempt -eq $maxRetries) {
                    Write-Host "  FAILED to rename $($file.Name) after $maxRetries attempts: $_" -ForegroundColor Red
                    Write-Log "ERROR: Failed to rename $($file.Name) after $maxRetries attempts: $_"
                    throw
                }
                Write-Host "  File locked: $($file.Name) - retrying in ${retryDelay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay
            }
        }
        $episodeNum++
    }
    Write-Host "Renamed $($episodeFiles.Count) episode(s)" -ForegroundColor Green
    Write-Log "Renamed $($episodeFiles.Count) episode(s) to Jellyfin format"
} else {
    # ========== MOVIE MODE: Original behavior ==========
    # prefix files with parent dir name (only if not already prefixed)
    # For disc 2+, add "Special Features-" after the movie name prefix
    if ($isMainFeatureDisc) {
        Write-Host "`nPrefixing files with directory name..." -ForegroundColor Yellow
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ($_.Directory.Name + "-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
            $filesToPrefix | ForEach-Object {
                $file = $_
                $newName = $file.Directory.Name + "-" + $file.Name
                $maxRetries = 5
                $retryDelay = 3
                for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                    try {
                        Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
                        break
                    } catch [System.IO.IOException] {
                        if ($attempt -eq $maxRetries) {
                            Write-Host "  FAILED to rename $($file.Name) after $maxRetries attempts: $_" -ForegroundColor Red
                            Write-Log "ERROR: Failed to rename $($file.Name) after $maxRetries attempts: $_"
                            throw
                        }
                        Write-Host "  File locked: $($file.Name) - retrying in ${retryDelay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
                        Start-Sleep -Seconds $retryDelay
                    }
                }
            }
            Write-Host "Prefixing complete" -ForegroundColor Green
            Write-Log "Prefixed $($filesToPrefix.Count) file(s) with directory name"
        } else {
            Write-Host "No files need prefixing" -ForegroundColor Gray
        }
    } else {
        # Disc 2+: prefix with "MovieName-Special Features-originalfilename"
        Write-Host "`nPrefixing special features files..." -ForegroundColor Yellow
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ($_.Directory.Name + "-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object {
                $file = $_
                $newName = $file.Directory.Name + "-Special Features-" + $file.Name
                Write-Host "  - $($file.Name) -> $newName" -ForegroundColor Gray
                $maxRetries = 5
                $retryDelay = 3
                for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                    try {
                        Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
                        break
                    } catch [System.IO.IOException] {
                        if ($attempt -eq $maxRetries) {
                            Write-Host "  FAILED to rename $($file.Name) after $maxRetries attempts: $_" -ForegroundColor Red
                            Write-Log "ERROR: Failed to rename $($file.Name) after $maxRetries attempts: $_"
                            throw
                        }
                        Write-Host "  File locked: $($file.Name) - retrying in ${retryDelay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
                        Start-Sleep -Seconds $retryDelay
                    }
                }
            }
            Write-Host "Special features prefixing complete" -ForegroundColor Green
            Write-Log "Prefixed $($filesToPrefix.Count) special features file(s)"
        } else {
            Write-Host "No files need prefixing" -ForegroundColor Gray
        }
    }

    # Movie disc 1 only: add 'Feature' suffix to largest file
    if ($isMainFeatureDisc) {
        Write-Host "`nChecking for Feature file..." -ForegroundColor Yellow
        $featureExists = Get-ChildItem -File | Where-Object { $_.Name -like "*-Feature.*" }
        if (!$featureExists) {
            $largestFile = Get-ChildItem -File | Sort-Object Length -Descending | Select-Object -First 1
            if ($largestFile) {
                Write-Host "Largest file: $($largestFile.Name) ($([math]::Round($largestFile.Length/1GB, 2)) GB)" -ForegroundColor White
                $newName = $largestFile.Directory.Name + "-Feature" + $largestFile.Extension
                Write-Host "Renaming to: $newName" -ForegroundColor Yellow
                $maxRetries = 5
                $retryDelay = 3
                for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                    try {
                        Rename-Item -LiteralPath $largestFile.FullName -NewName $newName -ErrorAction Stop
                        break
                    } catch [System.IO.IOException] {
                        if ($attempt -eq $maxRetries) {
                            Write-Host "  FAILED to rename $($largestFile.Name) after $maxRetries attempts: $_" -ForegroundColor Red
                            Write-Log "ERROR: Failed to rename $($largestFile.Name) after $maxRetries attempts: $_"
                            throw
                        }
                        Write-Host "  File locked: $($largestFile.Name) - retrying in ${retryDelay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
                        Start-Sleep -Seconds $retryDelay
                    }
                }
                Write-Host "Feature file renamed successfully" -ForegroundColor Green
                Write-Log "Feature file: $($largestFile.Name) -> $newName ($([math]::Round($largestFile.Length/1GB, 2)) GB)"
            }
        } else {
            Write-Host "Feature file already exists: $($featureExists.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "`nSkipping Feature rename (Special Features disc)" -ForegroundColor Gray
    }

    # Handle extras folder based on disc type
    if ($isMainFeatureDisc) {
        # Disc 1: move non-feature videos to extras
        Write-Host "`nChecking for non-feature videos..." -ForegroundColor Yellow
        $nonFeatureVideos = Get-ChildItem -File | Where-Object { $_.Extension -match '\.(mp4|avi|mkv|mov|wmv)$' -and $_.Name -notlike "*Feature*" }
        if ($nonFeatureVideos) {
            Write-Host "Non-feature videos found: $($nonFeatureVideos.Count)" -ForegroundColor White
            $nonFeatureVideos | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }

            if (!(Test-Path "extras")) {
                Write-Host "Creating extras directory..." -ForegroundColor Yellow
                md extras | Out-Null
                Write-Host "Extras directory created" -ForegroundColor Green
            } else {
                Write-Host "Extras directory already exists" -ForegroundColor Gray
            }

            Write-Host "Moving files to extras..." -ForegroundColor Yellow
            $nonFeatureVideos | Move-Item -Destination extras -ErrorAction SilentlyContinue
            Write-Host "Files moved to extras" -ForegroundColor Green
            Write-Log "Moved $($nonFeatureVideos.Count) non-feature file(s) to extras"
        } else {
            Write-Host "No non-feature videos found" -ForegroundColor Gray
        }
    } else {
        # Disc 2+: move videos to extras folder (exclude Feature file from disc 1)
        Write-Host "`nMoving special features to extras folder..." -ForegroundColor Yellow

        if (!(Test-Path $extrasDir)) {
            Write-Host "Creating extras directory..." -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $extrasDir | Out-Null
            Write-Host "Extras directory created" -ForegroundColor Green
        } else {
            Write-Host "Extras directory already exists" -ForegroundColor Gray
        }

        # Exclude Feature file (may have been created by disc 1)
        $videoFiles = Get-ChildItem -File | Where-Object { $_.Extension -match '\.(mp4|avi|mkv|mov|wmv)$' -and $_.Name -notlike "*-Feature.*" }
        if ($videoFiles) {
            Write-Host "Videos to move: $($videoFiles.Count)" -ForegroundColor White
            foreach ($video in $videoFiles) {
                $uniquePath = Get-UniqueFilePath -DestDir $extrasDir -FileName $video.Name
                $newName = [System.IO.Path]::GetFileName($uniquePath)
                if ($newName -ne $video.Name) {
                    Write-Host "  - $($video.Name) -> $newName (renamed to avoid clash)" -ForegroundColor Yellow
                } else {
                    Write-Host "  - $($video.Name)" -ForegroundColor Gray
                }
                Move-Item -Path $video.FullName -Destination $uniquePath
            }
            Write-Host "Files moved to extras" -ForegroundColor Green
            Write-Log "Moved $($videoFiles.Count) special features file(s) to extras"
        } else {
            Write-Host "No video files to move" -ForegroundColor Gray
        }
    }
}
Complete-CurrentStep
Write-Log "STEP 3/4: File organization complete"

# ========== STEP 4: OPEN DIRECTORY ==========
Set-CurrentStep -StepNumber 4
Write-Log "STEP 4/4: Opening directory..."
Write-Host "`n[STEP 4/4] Opening film directory..." -ForegroundColor Green
Write-Host "Opening: $finalOutputDir" -ForegroundColor Yellow
start $finalOutputDir
Complete-CurrentStep

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

# Always show title being processed
Write-Host "`nProcessed: $(Get-TitleSummary)" -ForegroundColor White
Write-Host "Final location: $finalOutputDir" -ForegroundColor White

# Show summary of completed steps
Show-StepsSummary

# File summary
Write-Host "`n--- FILE SUMMARY ---" -ForegroundColor Cyan
if ($script:EncodedFilesTooSmall) {
    Write-Host "  No large video files found" -ForegroundColor Red
    Write-Host "  Source MKV files preserved at: $makemkvOutputDir" -ForegroundColor Yellow
    Write-Host "  Log file: $($script:LogFile)" -ForegroundColor White
} else {
    $finalFiles = Get-ChildItem -Path $finalOutputDir -File -Recurse
    $totalSize = [math]::Round(($finalFiles | Measure-Object -Property Length -Sum).Sum/1GB, 2)
    Write-Host "  Total files: $($finalFiles.Count)" -ForegroundColor White
    Write-Host "  Total size: $totalSize GB" -ForegroundColor White
    Write-Host "  Log file: $($script:LogFile)" -ForegroundColor White
}
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Log "========== RIP SESSION COMPLETE =========="
Write-Log "Final location: $finalOutputDir"
if ($script:EncodedFilesTooSmall) {
    Write-Log "WARNING: Encoded files were too small - source MKV files preserved"
    Write-Log "Source MKV directory: $makemkvOutputDir"
} else {
    Write-Log "Total files: $($finalFiles.Count)"
    Write-Log "Total size: $totalSize GB"
    foreach ($f in $finalFiles) {
        Write-Log "  $($f.Name) ($([math]::Round($f.Length/1GB, 2)) GB)"
    }
}

Enable-ConsoleClose
$host.UI.RawUI.WindowTitle = "$windowTitle - DONE"
