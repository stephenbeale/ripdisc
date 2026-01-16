param(
    [Parameter(Mandatory=$true)]
    [string]$title,

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [switch]$Documentary,

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$Drive = "D:",

    [Parameter()]
    [int]$DriveIndex = -1,

    [Parameter()]
    [switch]$MultiPart  # For multi-disc movies with main content across discs
)

# ========== HELPER FUNCTIONS ==========
function Get-UniqueFilePath {
    param([string]$DestDir, [string]$FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $targetPath = Join-Path $DestDir $FileName

    if (!(Test-Path $targetPath)) { return $targetPath }

    $counter = 1
    do {
        $newName = "$baseName-$counter$extension"
        $targetPath = Join-Path $DestDir $newName
        $counter++
    } while (Test-Path $targetPath)

    return $targetPath
}

# Normalize drive letter format
$driveLetter = if ($Drive -match ':$') { $Drive } else { "${Drive}:" }

# ========== DRIVE CONFIRMATION ==========
$driveDescription = if ($DriveIndex -ge 0) { "Drive Index $DriveIndex" } else { "Drive $driveLetter" }
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Ready to rip: $title" -ForegroundColor White
Write-Host "Using: $driveDescription" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Read-Host "Press Enter to continue, or Ctrl+C to abort"

# ========== CONFIGURATION ==========
$makemkvOutputDir = "C:\Video\$title"  # temporary MakeMKV output
$finalOutputDir = if ($Series) { "F:\Series\$title\Season 01" } elseif ($Documentary) { "F:\Documentaries\$title" } else { "F:\DVDs\$title" }
$makemkvconPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"
$handbrakePath = "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe"

$lastSuccessfulStep = "None"

function Stop-WithError { param([string]$Step,[string]$Message)
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "FAILED!" -ForegroundColor Red
    Write-Host "Error at: $Step" -ForegroundColor Red
    Write-Host "Message: $Message" -ForegroundColor Red
    Write-Host "Last successful step: $lastSuccessfulStep" -ForegroundColor Yellow
    exit 1
}

$contentType = if ($Series) { "TV Series" } elseif ($Documentary) { "Documentary" } else { "Movie" }
$isMainFeatureDisc = (-not $Series -and -not $Documentary)

$extrasDir = Join-Path $finalOutputDir "extras"

# ========== STEP 1: RIP WITH MAKEMKV ==========
Write-Host "`n[STEP 1/4] Starting MakeMKV rip..." -ForegroundColor Green

if ($DriveIndex -ge 0) {
    # Determine actual drive letter from MakeMKV info
    $allDrives = & "$makemkvconPath" info
    $driveMap = $allDrives | Where-Object { $_ -match '^DRV:' } | ForEach-Object {
        $parts = $_ -split ','
        [PSCustomObject]@{ Index=[int]$parts[1]; Label=$parts[5].Trim('"'); Letter=$parts[6].Trim('"') }
    }
    $selectedDrive = $driveMap | Where-Object { $_.Index -eq $DriveIndex }
    if (-not $selectedDrive) { Stop-WithError "Drive selection" "DriveIndex $DriveIndex not found" }
    $discSource = "disc:$($selectedDrive.Index)"
    $ejectDrive = $selectedDrive.Letter
    Write-Host "Using drive index $DriveIndex ($ejectDrive)" -ForegroundColor Green
} else {
    $discSource = "dev:$driveLetter"
    $ejectDrive = $driveLetter
    Write-Host "Using drive: $driveLetter" -ForegroundColor Yellow
}

# Prepare temporary directory
if (Test-Path $makemkvOutputDir) {
    $existingMkvs = Get-ChildItem $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
    if ($existingMkvs) { $existingMkvs | Remove-Item -Force }
} else { New-Item -ItemType Directory -Path $makemkvOutputDir | Out-Null }

& $makemkvconPath mkv $discSource all $makemkvOutputDir --minlength=120
if ($LASTEXITCODE -ne 0) { Stop-WithError "MakeMKV rip" "Exited with code $LASTEXITCODE" }

$rippedFiles = Get-ChildItem $makemkvOutputDir -Filter "*.mkv"
if (!$rippedFiles) { Stop-WithError "MakeMKV rip" "No MKV files created." }

$lastSuccessfulStep = "STEP 1/4: MakeMKV rip"

# ========== EJECT DISC ==========
if ($ejectDrive -notmatch ':$') { $ejectDrive += ':' }
$cdDrive = Get-CimInstance Win32_CDROMDrive | Where-Object { $_.Drive -eq $ejectDrive }
if ($cdDrive) {
    try { $cdDrive.Eject() | Out-Null; Start-Sleep -Seconds 2; Write-Host "Disc ejected successfully via CIM" -ForegroundColor Green }
    catch {
        try { $shell = New-Object -ComObject Shell.Application; $shell.Namespace(17).ParseName($ejectDrive).InvokeVerb("Eject"); Write-Host "Disc ejected via COM fallback" -ForegroundColor Green } 
        catch { Write-Warning "Failed to eject disc on $ejectDrive" }
    }
} else { Write-Warning "No optical drive found for $ejectDrive" }

# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
Write-Host "`n[STEP 2/4] Starting HandBrake encoding..." -ForegroundColor Green
if (!(Test-Path $finalOutputDir)) { New-Item -ItemType Directory -Path $finalOutputDir | Out-Null }

$fileCount = 0
foreach ($mkv in $rippedFiles) {
    $fileCount++
    $inputFile = $mkv.FullName
    $outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")

    & $handbrakePath -i $inputFile -o $outputFile --preset "Fast 1080p30" --all-audio --all-subtitles --subtitle-burned=none --verbose=1
    if ($LASTEXITCODE -ne 0) { Stop-WithError "HandBrake encoding" "Failed encoding $($mkv.Name)" }
}
$lastSuccessfulStep = "STEP 2/4: HandBrake encoding"

# Clean up temporary MakeMKV folder
if (Get-ChildItem $finalOutputDir -Filter "*.mp4") { Remove-Item $makemkvOutputDir -Recurse -Force }

# ========== STEP 3: ORGANIZE FILES ==========
Write-Host "`n[STEP 3/4] Organizing files..." -ForegroundColor Green
cd $finalOutputDir

# --- Series ---
if ($Series) {
    # Determine next episode number across all discs
    $seriesRoot = Split-Path $finalOutputDir -Parent
    $seasonFolder = Split-Path $finalOutputDir -Leaf
    $seasonTag = "S01"  # hardcoded for now, can detect dynamically if needed

    $allEpisodes = Get-ChildItem $seriesRoot -Directory | ForEach-Object {
        $seasonPath = Join-Path $_.FullName $seasonFolder
        if (Test-Path $seasonPath) { Get-ChildItem $seasonPath -File -Filter "*.mp4" }
    } | ForEach-Object {
        if ($_.Name -match "$seasonTag`E(\d{2})") { [int]$Matches[1] }
    }

    $nextEpisode = if ($allEpisodes) { ($allEpisodes | Measure-Object -Maximum).Maximum + 1 } else { 1 }

    Get-ChildItem -File -Filter "*.mp4" | Sort-Object Name | ForEach-Object {
        $ep = "{0:D2}" -f $nextEpisode
        $newName = "$seasonTag`E$ep$($_.Extension)"
        Rename-Item $_ -NewName $newName
        $nextEpisode++
    }
}

# --- Movie / Documentary Extras ---
if (-not $Series -and -not $Documentary) {
    if (!(Test-Path $extrasDir)) { New-Item -ItemType Directory -Path $extrasDir | Out-Null }

    $videoFiles = Get-ChildItem -File | Where-Object { $_.Extension -match '\.(mp4|avi|mkv|mov|wmv)$' -and $_.Name -notlike "*-Feature.*" }
    foreach ($video in $videoFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
        if ($baseName -notlike "$title*") { $newName = "$title-$($video.Name)" } else { $newName = $video.Name }
        $uniquePath = Get-UniqueFilePath -DestDir $extrasDir -FileName $newName
        Move-Item $video.FullName $uniquePath
    }
}

# ========== STEP 4: OPEN DIRECTORY ==========
Write-Host "`n[STEP 4/4] Opening output directory..." -ForegroundColor Green
Start $finalOutputDir

$lastSuccessfulStep = "STEP 4/4: Complete"
Write-Host "`nAll steps completed successfully." -ForegroundColor Green
