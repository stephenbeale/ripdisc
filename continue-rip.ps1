param(
    [Parameter(Mandatory=$true)]
    [string]$title,

    [Parameter(Mandatory=$true)]
    [ValidateSet("handbrake", "organize", "open")]
    [string]$FromStep,

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [int]$Season = 0,

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$OutputDrive = "E:",

    [Parameter()]
    [switch]$Extras,

    [Parameter()]
    [switch]$Bluray,

    [Parameter()]
    [switch]$Documentary
)

# ========== STEP MAPPING ==========
$StepMapping = @{
    "handbrake" = 2
    "organize" = 3
    "open" = 4
}
$StartFromStepNumber = $StepMapping[$FromStep]

# ========== STEP TRACKING ==========
$script:AllSteps = @(
    @{ Number = 1; Name = "MakeMKV rip"; Description = "Rip disc to MKV files" }
    @{ Number = 2; Name = "HandBrake encoding"; Description = "Encode MKV to MP4" }
    @{ Number = 3; Name = "Organize files"; Description = "Rename and move files" }
    @{ Number = 4; Name = "Open directory"; Description = "Open output folder" }
)
$script:CompletedSteps = @()
$script:CurrentStep = $null
$script:LastWorkingDirectory = $null

# Mark steps before starting point as "skipped/assumed complete"
for ($i = 1; $i -lt $StartFromStepNumber; $i++) {
    $script:CompletedSteps += $script:AllSteps | Where-Object { $_.Number -eq $i }
}

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
    [Win32.ConsoleCloseProtection]::EnableMenuItem($script:ConsoleSystemMenu, 0xF060, 0x00000001) | Out-Null
}

function Enable-ConsoleClose {
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

    $driveLetter = [System.IO.Path]::GetPathRoot($Path)
    if (-not $driveLetter) {
        return @{ Ready = $false; Drive = "Unknown"; Message = "Could not determine drive letter from path: $Path" }
    }

    $driveDisplay = $driveLetter.TrimEnd('\')

    try {
        $drive = Get-PSDrive -Name $driveDisplay.TrimEnd(':') -ErrorAction Stop
        if ($drive) {
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

# ========== CONFIGURATION ==========
# MakeMKV temp directory - use subdirectory for multi-disc and extras rips
if ($Extras) {
    $makemkvOutputDir = "C:\Video\$title\Extras"
} else {
    $makemkvOutputDir = "C:\Video\$title\Disc$Disc"
}

# Normalize output drive letter
$outputDriveLetter = if ($OutputDrive -match ':$') { $OutputDrive } else { "${OutputDrive}:" }

# Build final output directory path
if ($Documentary) {
    $finalOutputDir = "$outputDriveLetter\Documentaries\$title"
} elseif ($Series) {
    $seriesBaseDir = "$outputDriveLetter\Series\$title"
    if ($Season -gt 0) {
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $finalOutputDir = Join-Path $seriesBaseDir $seasonFolder
    } else {
        $seasonTag = $null
        $finalOutputDir = $seriesBaseDir
    }
} else {
    $finalOutputDir = "$outputDriveLetter\DVDs\$title"
}

$handbrakePath = "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe"

# ========== LOGGING SETUP ==========
$logDir = "C:\Video\logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDiscLabel = if ($Extras) { "extras" } else { "disc${Disc}" }
$script:LogFile = Join-Path $logDir "${title}_${logDiscLabel}_continue_${logTimestamp}.log"

Write-Log "========== CONTINUE SESSION STARTED =========="
Write-Log "Title: $title"
Write-Log "Continue from: Step $StartFromStepNumber ($FromStep)"
Write-Log "Type: $(if ($Documentary) { 'Documentary' } elseif ($Series) { 'TV Series' } else { 'Movie' })"
Write-Log "Disc: $Disc$(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1 -and -not $Series) { ' (Special Features)' })"
if ($Series -and $Season -gt 0) {
    Write-Log "Season: $Season"
}
Write-Log "Output Drive: $outputDriveLetter"
Write-Log "MakeMKV Output: $makemkvOutputDir"
Write-Log "Final Output: $finalOutputDir"
Write-Log "Log file: $($script:LogFile)"

function Stop-WithError {
    param([string]$Step, [string]$Message)

    $host.UI.RawUI.WindowTitle = "$($host.UI.RawUI.WindowTitle) - ERROR"

    Write-Log "========== ERROR =========="
    Write-Log "Failed at: $Step"
    Write-Log "Message: $Message"
    if ($script:CompletedSteps.Count -gt 0) {
        Write-Log "Completed steps: $(($script:CompletedSteps | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    }
    $remaining = Get-RemainingSteps
    if ($remaining.Count -gt 0) {
        Write-Log "Remaining steps: $(($remaining | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    }

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "`nProcessing: $(Get-TitleSummary)" -ForegroundColor White
    Write-Host "`nError at: $Step" -ForegroundColor Red
    Write-Host "Message: $Message" -ForegroundColor Red

    Show-StepsSummary -ShowRemaining

    $directoryToOpen = $null
    if ($script:LastWorkingDirectory -and (Test-Path $script:LastWorkingDirectory)) {
        $directoryToOpen = $script:LastWorkingDirectory
    } elseif (Test-Path $makemkvOutputDir) {
        $directoryToOpen = $makemkvOutputDir
    } elseif (Test-Path $finalOutputDir) {
        $directoryToOpen = $finalOutputDir
    }

    if ($directoryToOpen) {
        Write-Host "`n--- OPENING DIRECTORY ---" -ForegroundColor Cyan
        Write-Host "Opening: $directoryToOpen" -ForegroundColor Yellow
        Start-Process explorer.exe -ArgumentList $directoryToOpen
    }

    Write-Host "`nLog file: $($script:LogFile)" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Red
    Enable-ConsoleClose
    exit 1
}

$contentType = if ($Documentary) { "Documentary" } elseif ($Series) { "TV Series" } else { "Movie" }
$isMainFeatureDisc = (-not $Series) -and ($Disc -eq 1) -and (-not $Extras)
$extrasDir = Join-Path $finalOutputDir "extras"

# ========== WINDOW TITLE ==========
if ($Series) {
    $windowTitle = "$title"
    if ($Season -gt 0) { $windowTitle += " S$Season" }
    $windowTitle += " Disc $Disc"
} else {
    $windowTitle = "$title"
    if ($Extras -or $Disc -gt 1) { $windowTitle += "-extras" }
}
$windowTitle += " - CONTINUE"
$host.UI.RawUI.WindowTitle = $windowTitle

# ========== VALIDATION ==========
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Continue Rip - Starting from $FromStep" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Title: $title" -ForegroundColor White
Write-Host "Type: $contentType" -ForegroundColor White
Write-Host "Output Drive: $outputDriveLetter" -ForegroundColor White
if ($Series) {
    if ($Season -gt 0) {
        Write-Host "Season: $Season ($seasonTag)" -ForegroundColor White
    }
    Write-Host "Disc: $Disc" -ForegroundColor White
} else {
    Write-Host "Disc: $Disc$(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1) { ' (Special Features)' })" -ForegroundColor White
}
Write-Host "MakeMKV Output: $makemkvOutputDir" -ForegroundColor White
Write-Host "Final Output: $finalOutputDir" -ForegroundColor White
Write-Host "Starting from: Step $StartFromStepNumber ($FromStep)" -ForegroundColor Yellow
Write-Host "Log file: $($script:LogFile)" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

# Validate prerequisites for starting step
if ($StartFromStepNumber -eq 2) {
    # Need MKV files in makemkvOutputDir
    if (!(Test-Path $makemkvOutputDir)) {
        Write-Host "ERROR: MakeMKV output directory not found: $makemkvOutputDir" -ForegroundColor Red
        Write-Host "Cannot continue from HandBrake step without MKV files." -ForegroundColor Red
        exit 1
    }
    $mkvFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
    if ($null -eq $mkvFiles -or $mkvFiles.Count -eq 0) {
        Write-Host "ERROR: No MKV files found in: $makemkvOutputDir" -ForegroundColor Red
        Write-Host "Cannot continue from HandBrake step without MKV files." -ForegroundColor Red
        exit 1
    }
    Write-Host "Found $($mkvFiles.Count) MKV file(s) to encode:" -ForegroundColor Green
    foreach ($mkv in $mkvFiles) {
        Write-Host "  - $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB)" -ForegroundColor Gray
    }
} elseif ($StartFromStepNumber -eq 3) {
    # Need MP4 files in finalOutputDir
    if (!(Test-Path $finalOutputDir)) {
        Write-Host "ERROR: Final output directory not found: $finalOutputDir" -ForegroundColor Red
        Write-Host "Cannot continue from organize step without encoded files." -ForegroundColor Red
        exit 1
    }
    $mp4Files = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4" -ErrorAction SilentlyContinue
    if ($null -eq $mp4Files -or $mp4Files.Count -eq 0) {
        Write-Host "ERROR: No MP4 files found in: $finalOutputDir" -ForegroundColor Red
        Write-Host "Cannot continue from organize step without encoded files." -ForegroundColor Red
        exit 1
    }
    Write-Host "Found $($mp4Files.Count) MP4 file(s) to organize:" -ForegroundColor Green
    foreach ($mp4 in $mp4Files) {
        Write-Host "  - $($mp4.Name) ($([math]::Round($mp4.Length/1GB, 2)) GB)" -ForegroundColor Gray
    }
} elseif ($StartFromStepNumber -eq 4) {
    # Just need finalOutputDir to exist
    if (!(Test-Path $finalOutputDir)) {
        Write-Host "ERROR: Final output directory not found: $finalOutputDir" -ForegroundColor Red
        exit 1
    }
    Write-Host "Output directory exists: $finalOutputDir" -ForegroundColor Green
}

$response = Read-Host "`nPress Enter to continue, or Ctrl+C to abort"
Disable-ConsoleClose

# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
if ($StartFromStepNumber -le 2) {
    Set-CurrentStep -StepNumber 2
    $script:LastWorkingDirectory = $finalOutputDir
    Write-Log "STEP 2/4: Starting HandBrake encoding..."
    Write-Host "`n[STEP 2/4] Starting HandBrake encoding..." -ForegroundColor Green

    # Check if destination drive is ready
    Write-Host "Checking destination drive..." -ForegroundColor Yellow
    $driveCheck = Test-DriveReady -Path $finalOutputDir
    if (-not $driveCheck.Ready) {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message $driveCheck.Message
    }
    Write-Host "Destination drive $($driveCheck.Drive) is ready" -ForegroundColor Green

    Write-Host "Creating directory: $finalOutputDir" -ForegroundColor Yellow
    if (!(Test-Path $finalOutputDir)) {
        New-Item -ItemType Directory -Path $finalOutputDir | Out-Null
        Write-Host "Directory created successfully" -ForegroundColor Green
    } else {
        Write-Host "Directory already exists" -ForegroundColor Yellow
    }

    $mkvFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv"
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

    # Wait for file handles
    Write-Host "`nWaiting for file handles to be released..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    Write-Host "File handle wait complete" -ForegroundColor Green

    # Delete MakeMKV temp directory after successful encode
    Write-Host "`nChecking for successful encodes..." -ForegroundColor Yellow
    $encodedFiles = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4"
    $script:EncodedFilesTooSmall = $false
    if ($encodedFiles.Count -gt 0) {
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
            Start-Process explorer.exe -ArgumentList "shell:RecycleBinFolder"
            Write-Host "Opened Recycle Bin for review" -ForegroundColor Yellow
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
}

# ========== STEP 3: RENAME AND ORGANIZE ==========
if ($StartFromStepNumber -le 3) {
    Set-CurrentStep -StepNumber 3
    $script:LastWorkingDirectory = $finalOutputDir
    Write-Log "STEP 3/4: Organizing files..."
    Write-Host "`n[STEP 3/4] Organizing files..." -ForegroundColor Green
    cd $finalOutputDir

    Write-Host "Current directory: $finalOutputDir" -ForegroundColor Yellow

    # Delete image files
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
        # ========== SERIES MODE ==========
        Write-Host "`nPrefixing files with title..." -ForegroundColor Yellow
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike "$title-*" }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
            $filesToPrefix | ForEach-Object {
                $file = $_
                $newName = "$title-" + $file.Name
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
            Write-Log "Prefixed $($filesToPrefix.Count) file(s) with title"
        } else {
            Write-Host "No files need prefixing" -ForegroundColor Gray
        }
    } else {
        # ========== MOVIE MODE ==========
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

        # Handle extras folder
        if ($isMainFeatureDisc) {
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
            Write-Host "`nMoving special features to extras folder..." -ForegroundColor Yellow

            if (!(Test-Path $extrasDir)) {
                Write-Host "Creating extras directory..." -ForegroundColor Yellow
                New-Item -ItemType Directory -Path $extrasDir | Out-Null
                Write-Host "Extras directory created" -ForegroundColor Green
            } else {
                Write-Host "Extras directory already exists" -ForegroundColor Gray
            }

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
}

# ========== STEP 4: OPEN DIRECTORY ==========
if ($StartFromStepNumber -le 4) {
    Set-CurrentStep -StepNumber 4
    Write-Log "STEP 4/4: Opening directory..."
    Write-Host "`n[STEP 4/4] Opening film directory..." -ForegroundColor Green
    Write-Host "Opening: $finalOutputDir" -ForegroundColor Yellow
    start $finalOutputDir
    Complete-CurrentStep
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nProcessed: $(Get-TitleSummary)" -ForegroundColor White
Write-Host "Final location: $finalOutputDir" -ForegroundColor White

Show-StepsSummary

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

Write-Log "========== CONTINUE SESSION COMPLETE =========="
Write-Log "Final location: $finalOutputDir"
if (-not $script:EncodedFilesTooSmall) {
    Write-Log "Total files: $($finalFiles.Count)"
    Write-Log "Total size: $totalSize GB"
    foreach ($f in $finalFiles) {
        Write-Log "  $($f.Name) ($([math]::Round($f.Length/1GB, 2)) GB)"
    }
}

Enable-ConsoleClose
$host.UI.RawUI.WindowTitle = "$windowTitle - DONE"
