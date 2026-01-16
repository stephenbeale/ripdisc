param(
    [Parameter(Mandatory=$true)]
    [string]$title,

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [switch]$Documentary,

    [Parameter()]
    [switch]$MultiPart,  # for multi-disc movies like Dances with Wolves

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$Drive = "D:",

    [Parameter()]
    [int]$DriveIndex = -1
)

# -------- Helper: Get unique file path to avoid overwrites --------
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

# -------- Normalize drive --------
$driveLetter = if ($Drive -match ':$') { $Drive } else { "${Drive}:" }

# -------- Show drive info and confirm --------
$driveDescription = if ($DriveIndex -ge 0) {
    switch ($DriveIndex) {
        0 { "D: Black Sandstrom" }
        1 { "G: White Sandstrom" }
        default { "unknown drive" }
    }
} else { "Drive $driveLetter" }

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Ready to rip: $title" -ForegroundColor White
Write-Host "Using: $driveDescription" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Read-Host "Press Enter to continue, or Ctrl+C to abort"

# -------- Paths --------
$makemkvOutputDir = "C:\Video\$title"
$finalOutputDir  = if ($Series) { "F:\Series\$title\Season 01" } elseif ($Documentary) { "F:\Documentaries\$title"} else { "F:\DVDs\$title" }
$extrasDir       = Join-Path $finalOutputDir "extras"
$makemkvconPath  = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"
$handbrakePath   = "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe"

$lastSuccessfulStep = "None"

function Stop-WithError {
    param([string]$Step, [string]$Message)
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "FAILED!" -ForegroundColor Red
    Write-Host "Error at: $Step" -ForegroundColor Red
    Write-Host "Message: $Message" -ForegroundColor Red
    Write-Host "Last successful step: $lastSuccessfulStep" -ForegroundColor Yellow
    exit 1
}

$contentType = if ($Series) { "TV Series" } elseif ($Documentary) { "Documentary" } else { "Movie" }
$isMainFeatureDisc = (-not $Series -and -not $Documentary) -and ($Disc -eq 1)

# -------- MakeMKV Rip --------
Write-Host "`n[STEP 1/4] Starting MakeMKV rip..." -ForegroundColor Green

if ($DriveIndex -ge 0) { $discSource = "disc:$DriveIndex" } else { $discSource = "dev:$driveLetter" }

if (!(Test-Path $makemkvOutputDir)) { New-Item -ItemType Directory -Path $makemkvOutputDir | Out-Null }

& $makemkvconPath mkv $discSource all $makemkvOutputDir --minlength=120
if ($LASTEXITCODE -ne 0) { Stop-WithError "MakeMKV rip" "Exit code $LASTEXITCODE" }

$rippedFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv"
if (!$rippedFiles) { Stop-WithError "MakeMKV rip" "No MKV files created" }
$lastSuccessfulStep = "STEP 1/4: MakeMKV rip"

# -------- Eject Disc (CIM + COM fallback) --------
$ejectDrive = if ($DriveIndex -ge 0) {
    switch ($DriveIndex) {0 {"D:"} 1 {"G:"} default {$driveLetter}}
} else { $driveLetter }
if ($ejectDrive -notmatch ':$') { $ejectDrive += ':' }

Write-Host "`nEjecting disc $ejectDrive..." -ForegroundColor Yellow
$cdDrive = Get-CimInstance Win32_CDROMDrive | Where-Object { $_.Drive -eq $ejectDrive }
if ($cdDrive) {
    try { $cdDrive.Eject() | Out-Null; Start-Sleep -Seconds 2; Write-Host "Disc ejected successfully" -ForegroundColor Green }
    catch { $shell = New-Object -ComObject Shell.Application; $shell.Namespace(17).ParseName($ejectDrive).InvokeVerb("Eject"); Write-Host "Disc ejected via COM fallback" -ForegroundColor Green }
} else { Write-Warning "No optical drive found for $ejectDrive" }

# -------- HandBrake Encode --------
Write-Host "`n[STEP 2/4] Starting HandBrake encoding..." -ForegroundColor Green
if (!(Test-Path $finalOutputDir)) { New-Item -ItemType Directory -Path $finalOutputDir | Out-Null }

$encodedCount = 0
foreach ($mkv in $rippedFiles) {
    $encodedCount++
    $outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")
    & $handbrakePath -i $mkv.FullName -o $outputFile --preset "Fast 1080p30" --all-audio --all-subtitles --subtitle-burned=none --verbose=1
    if ($LASTEXITCODE -ne 0) { Stop-WithError "HandBrake encoding" "Exit code $LASTEXITCODE on $($mkv.Name)" }
    Write-Host "Encoding complete: $($mkv.Name) -> $outputFile" -ForegroundColor Green
}
$lastSuccessfulStep = "STEP 2/4: HandBrake encoding"

# Remove temp MKV dir
if (Test-Path $makemkvOutputDir) { Remove-Item -Path $makemkvOutputDir -Recurse -Force }

# -------- File Organization --------
Write-Host "`n[STEP 3/4] Organizing files..." -ForegroundColor Green
Set-Location $finalOutputDir

# -------- Series renaming --------
if ($Series) {
    # Detect season folder
    $seasonFolder = Get-ChildItem -Directory | Select-Object -First 1
    $seasonNumber = 1
    if ($seasonFolder -and $seasonFolder.Name -match '\d+') { $seasonNumber = [int]$Matches[0] }
    $seasonTag = "S{0:D2}" -f $seasonNumber
    if (-not $seasonFolder -or $seasonFolder.Name -ne "Season $seasonNumber") { $seasonFolder = New-Item -ItemType Directory -Name "Season $seasonNumber" -Force }

    # Determine next episode number from all discs
    $seriesParent = Split-Path $finalOutputDir -Parent
    $allSeasonEpisodes =
        Get-ChildItem $seriesParent -Directory |
        ForEach-Object {
            $path = Join-Path $_.FullName $seasonFolder.Name
            if (Test-Path $path) { Get-ChildItem $path -File -Filter "*.mp4" }
        } |
        Where-Object { $_.Name -match "$seasonTag`E(\d{2})" } |
        ForEach-Object { [int]$Matches[1] }

    $nextEpisode = if ($allSeasonEpisodes) { ($allSeasonEpisodes | Measure-Object -Maximum).Maximum + 1 } else { 1 }

    Get-ChildItem -File -Filter "*.mp4" |
        Sort-Object Name |
        ForEach-Object {
            $ep = "{0:D2}" -f $nextEpisode
            $newName = "$seasonTag`E$ep$($_.Extension)"
            Rename-Item $_ -NewName $newName
            $nextEpisode++
        }
}

# -------- Movie / MultiPart handling --------
elseif ($MultiPart -or -not $Series -and -not $Documentary) {
    $videoFiles = Get-ChildItem -File -Filter "*.mp4" | Sort-Object Length -Descending
    if ($videoFiles.Count -gt 0) {
        # Determine main feature file
        $featureFile = $videoFiles[0]
        $featureName = if ($MultiPart) { "$title - Part $Disc$($featureFile.Extension)" } else { "$title-Feature$($featureFile.Extension)" }
        Rename-Item $featureFile.FullName $featureName -Force

        # Move remaining files to extras
        $otherFiles = $videoFiles | Where-Object { $_.FullName -ne $featureFile.FullName }
        foreach ($file in $otherFiles) {
            $newName = if ($file.BaseName -notlike "$title*") { "$title-$($file.Name)" } else { $file.Name }
            $uniquePath = Get-UniqueFilePath -DestDir $extrasDir -FileName $newName
            if (!(Test-Path $extrasDir)) { New-Item -ItemType Directory -Path $extrasDir | Out-Null }
            Move-Item $file.FullName $uniquePath
        }
    }
}

# -------- Delete image files --------
Get-ChildItem -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png|gif|bmp)$' } | Remove-Item -Force -ErrorAction SilentlyContinue

$lastSuccessfulStep = "STEP 3/4: File organization"

# -------- Open final folder --------
Write-Host "`n[STEP 4/4] Opening final directory..." -ForegroundColor Green
Start-Process $finalOutputDir
$lastSuccessfulStep = "STEP 4/4: Open directory"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "COMPLETE!" -ForegroundColor Green
Write-Host "All steps completed successfully" -ForegroundColor Cyan
