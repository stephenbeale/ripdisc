param(
    [Parameter(Mandatory=$true)]
    [string]$title,

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [int]$Season = 0,

    [Parameter()]
    [int]$StartEpisode = 1,

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$Drive = "D:",

    [Parameter()]
    [int]$DriveIndex = -1,

    [Parameter()]
    [string]$OutputDrive = "E:"
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
    $contentType = if ($Series) { "TV Series" } else { "Movie" }
    $summary = "$contentType`: $title"
    if ($Series) {
        if ($Season -gt 0) {
            $summary += " - Season $Season, Disc $Disc"
        } else {
            $summary += " - Disc $Disc"
        }
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
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Ready to rip: $title" -ForegroundColor White
if ($Series) {
    if ($Season -gt 0) {
        $seasonTagPreview = "S{0:D2}" -f $Season
        Write-Host "Type: TV Series - Season $Season ($seasonTagPreview), Disc $Disc" -ForegroundColor White
        Write-Host "Starting at: E$("{0:D2}" -f $StartEpisode)" -ForegroundColor White
    } else {
        Write-Host "Type: TV Series - Disc $Disc (no season folder)" -ForegroundColor White
        Write-Host "Starting at: E$("{0:D2}" -f $StartEpisode)" -ForegroundColor White
    }
} else {
    $discType = if ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: Movie - $discType (Disc $Disc)" -ForegroundColor White
}
Write-Host "Using: $driveDescription" -ForegroundColor Yellow
Write-Host "Output Drive: $OutputDrive" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
$response = Read-Host "Press Enter to continue, or Ctrl+C to abort"

# ========== CONFIGURATION ==========
# Make MakeMKV temp directory disc-specific to support concurrent ripping
$makemkvOutputDir = "C:\Video\$title\Disc$Disc"  # MakeMKV rips here first

# Normalize output drive letter (add colon if missing)
$outputDriveLetter = if ($OutputDrive -match ':$') { $OutputDrive } else { "${OutputDrive}:" }

# Series: organize into Season subfolders with proper naming (only if Season explicitly specified)
# Movies: organize into title folder with optional extras
if ($Series) {
    $seriesBaseDir = "$outputDriveLetter\Series\$title"
    if ($Season -gt 0) {
        # Season explicitly specified - use Season subfolder and episode tags
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

function Stop-WithError {
    param([string]$Step, [string]$Message)

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
                    if ($seasonTag) {
                        Write-Host "    Format: $title - $seasonTag`E##.mp4" -ForegroundColor Gray
                    } else {
                        Write-Host "    Format: $title - E##.mp4" -ForegroundColor Gray
                    }
                } else {
                    if ($Disc -eq 1) {
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

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "Please complete the remaining steps manually" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    exit 1
}

$contentType = if ($Series) { "TV Series" } else { "Movie" }
$isMainFeatureDisc = (-not $Series) -and ($Disc -eq 1)
$extrasDir = Join-Path $finalOutputDir "extras"

# For disc 2+, ensure parent dir and extras folder exist upfront (disc 1 may still be running)
if (-not $isMainFeatureDisc -and -not $Series) {
    # Check if destination drive is ready before attempting to create directories
    $driveCheck = Test-DriveReady -Path $finalOutputDir
    if (-not $driveCheck.Ready) {
        Write-Host "`nERROR: $($driveCheck.Message)" -ForegroundColor Red
        exit 1
    }
    if (!(Test-Path $finalOutputDir)) {
        New-Item -ItemType Directory -Path $finalOutputDir -Force | Out-Null
    }
    if (!(Test-Path $extrasDir)) {
        New-Item -ItemType Directory -Path $extrasDir -Force | Out-Null
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
    Write-Host "Starting Episode: E$("{0:D2}" -f $StartEpisode)" -ForegroundColor White
} else {
    Write-Host "Disc: $Disc$(if ($Disc -gt 1) { ' (Special Features)' })" -ForegroundColor White
}
Write-Host "MakeMKV Output: $makemkvOutputDir" -ForegroundColor White
Write-Host "Final Output: $finalOutputDir" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

# ========== STEP 1: RIP WITH MAKEMKV ==========
Set-CurrentStep -StepNumber 1
$script:LastWorkingDirectory = $makemkvOutputDir
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
    Write-Host "Directory exists - cleaning existing MKV files..." -ForegroundColor Yellow
    $existingMkvs = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
    if ($existingMkvs) {
        $existingMkvs | Remove-Item -Force
        Write-Host "Removed $($existingMkvs.Count) existing MKV file(s)" -ForegroundColor Yellow
    }
} else {
    New-Item -ItemType Directory -Path $makemkvOutputDir | Out-Null
    Write-Host "Directory created successfully" -ForegroundColor Green
}

Write-Host "`nExecuting MakeMKV command..." -ForegroundColor Yellow
Write-Host "Command: makemkvcon mkv $discSource all `"$makemkvOutputDir`" --minlength=120" -ForegroundColor Gray

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
foreach ($file in $rippedFiles) {
    Write-Host "  - $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)" -ForegroundColor Gray
}
Complete-CurrentStep

# Eject disc
Write-Host "`nEjecting disc from drive $driveLetter..." -ForegroundColor Yellow
$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).ParseName($driveLetter).InvokeVerb("Eject")
Write-Host "Disc ejected successfully" -ForegroundColor Green


# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
Set-CurrentStep -StepNumber 2
$script:LastWorkingDirectory = $finalOutputDir
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

    Write-Host "`nExecuting HandBrake..." -ForegroundColor Yellow
    & $handbrakePath -i $inputFile -o $outputFile `
        --preset "Fast 1080p30" `
        --all-audio `
        --all-subtitles `
        --subtitle-burned=none `
        --verbose=1
    $handbrakeExitCode = $LASTEXITCODE

    if ($handbrakeExitCode -ne 0) {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "HandBrake exited with code $handbrakeExitCode while encoding $($mkv.Name)"
    }

    if (Test-Path $outputFile) {
        $encodedSize = (Get-Item $outputFile).Length
        Write-Host "`nEncoding complete: $($mkv.Name)" -ForegroundColor Green
        Write-Host "Output size: $([math]::Round($encodedSize/1GB, 2)) GB" -ForegroundColor White
    } else {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "Output file not created for $($mkv.Name)"
    }
}
Complete-CurrentStep

# Wait for HandBrake to fully release file handles before proceeding
Write-Host "`nWaiting for file handles to be released..." -ForegroundColor Yellow
Start-Sleep -Seconds 3
Write-Host "File handle wait complete" -ForegroundColor Green

# Delete MakeMKV temporary directory after successful encode
Write-Host "`nChecking for successful encodes..." -ForegroundColor Yellow
$encodedFiles = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4"
if ($encodedFiles.Count -gt 0) {
    Write-Host "Found $($encodedFiles.Count) encoded file(s)" -ForegroundColor Green
    Write-Host "Removing temporary MakeMKV directory: $makemkvOutputDir" -ForegroundColor Yellow
    Remove-Item -Path $makemkvOutputDir -Recurse -Force
    Write-Host "Temporary files removed successfully" -ForegroundColor Green
} else {
    Write-Host "WARNING: No encoded files found. Keeping MakeMKV directory." -ForegroundColor Red
}

# ========== STEP 3: RENAME AND ORGANIZE ==========
Set-CurrentStep -StepNumber 3
$script:LastWorkingDirectory = $finalOutputDir
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
    # ========== SERIES MODE: Episode naming ==========
    if ($seasonTag) {
        Write-Host "`nRenaming episodes with $seasonTag format..." -ForegroundColor Yellow
    } else {
        Write-Host "`nRenaming episodes (no season tag)..." -ForegroundColor Yellow
    }

    # Get all MP4 files sorted by name (MakeMKV typically names them title00.mkv, title01.mkv, etc.)
    $episodeFiles = Get-ChildItem -File -Filter "*.mp4" | Sort-Object Name

    if ($episodeFiles) {
        $currentEpisode = $StartEpisode
        Write-Host "Found $($episodeFiles.Count) episode(s) to rename" -ForegroundColor White
        Write-Host "Starting from episode E$("{0:D2}" -f $currentEpisode)" -ForegroundColor White

        foreach ($episode in $episodeFiles) {
            $episodeTag = "E{0:D2}" -f $currentEpisode
            if ($seasonTag) {
                $newName = "$title - $seasonTag$episodeTag$($episode.Extension)"
            } else {
                $newName = "$title - $episodeTag$($episode.Extension)"
            }

            Write-Host "  $($episode.Name) -> $newName" -ForegroundColor Yellow
            Rename-Item -Path $episode.FullName -NewName $newName
            $currentEpisode++
        }

        $lastEpisode = $currentEpisode - 1
        if ($seasonTag) {
            Write-Host "Episode renaming complete ($seasonTag`E$("{0:D2}" -f $StartEpisode) to $seasonTag`E$("{0:D2}" -f $lastEpisode))" -ForegroundColor Green
        } else {
            Write-Host "Episode renaming complete (E$("{0:D2}" -f $StartEpisode) to E$("{0:D2}" -f $lastEpisode))" -ForegroundColor Green
        }

        # Calculate and display next episode number for user convenience
        Write-Host "`n--- NEXT DISC INFO ---" -ForegroundColor Cyan
        Write-Host "For the next disc, use:" -ForegroundColor White
        Write-Host "  -StartEpisode $currentEpisode" -ForegroundColor Yellow
        Write-Host "----------------------" -ForegroundColor Cyan
    } else {
        Write-Host "No MP4 files found to rename" -ForegroundColor Gray
    }
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
            $filesToPrefix | Rename-Item -NewName { $_.Directory.Name + "-" + $_.Name }
            Write-Host "Prefixing complete" -ForegroundColor Green
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
                $newName = $_.Directory.Name + "-Special Features-" + $_.Name
                Write-Host "  - $($_.Name) -> $newName" -ForegroundColor Gray
                Rename-Item -Path $_.FullName -NewName $newName
            }
            Write-Host "Special features prefixing complete" -ForegroundColor Green
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
                Rename-Item -Path $largestFile.FullName -NewName $newName
                Write-Host "Feature file renamed successfully" -ForegroundColor Green
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
        } else {
            Write-Host "No video files to move" -ForegroundColor Gray
        }
    }
}
Complete-CurrentStep

# ========== STEP 4: OPEN DIRECTORY ==========
Set-CurrentStep -StepNumber 4
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
$finalFiles = Get-ChildItem -Path $finalOutputDir -File -Recurse
Write-Host "  Total files: $($finalFiles.Count)" -ForegroundColor White
Write-Host "  Total size: $([math]::Round(($finalFiles | Measure-Object -Property Length -Sum).Sum/1GB, 2)) GB" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan
