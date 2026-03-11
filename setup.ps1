<#
.SYNOPSIS
    First-run setup for RipDisc. Detects or installs MakeMKV and HandBrake,
    then writes a ripdisc-config.json so the ripping scripts work out of the box.
#>

$configPath = Join-Path $PSScriptRoot "ripdisc-config.json"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  RipDisc Setup" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ========== HELPER: FIND EXECUTABLE ==========
function Find-Executable {
    param(
        [string]$Name,
        [string[]]$SearchPaths,
        [string]$RegistryKey,
        [string]$RegistryValue = "InstallLocation"
    )

    # 1. Check PATH
    $inPath = Get-Command $Name -ErrorAction SilentlyContinue
    if ($inPath) { return $inPath.Source }

    # 2. Check registry install location
    if ($RegistryKey) {
        foreach ($root in @("HKLM:\SOFTWARE", "HKLM:\SOFTWARE\WOW6432Node", "HKCU:\SOFTWARE")) {
            $fullKey = Join-Path $root $RegistryKey
            $reg = Get-ItemProperty -Path $fullKey -ErrorAction SilentlyContinue
            if ($reg -and $reg.$RegistryValue) {
                $candidate = Join-Path $reg.$RegistryValue $Name
                if (Test-Path $candidate) { return $candidate }
            }
        }
    }

    # 3. Check common install paths
    foreach ($p in $SearchPaths) {
        if (Test-Path $p) { return $p }
    }

    return $null
}

# ========== HELPER: INSTALL VIA CHOCOLATEY ==========
function Install-ViaChocolatey {
    param([string]$PackageName, [string]$DisplayName)

    $choco = Get-Command choco -ErrorAction SilentlyContinue
    if (-not $choco) {
        Write-Host ""
        Write-Host "  Chocolatey is not installed yet." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Chocolatey is a package manager for Windows (like apt or brew)." -ForegroundColor Gray
        Write-Host "  It lets you install and update software from the command line." -ForegroundColor Gray
        Write-Host ""
        Write-Host "  To install Chocolatey:" -ForegroundColor White
        Write-Host "    1. Open PowerShell as Administrator (right-click > Run as Administrator)" -ForegroundColor Gray
        Write-Host "    2. Run this command:" -ForegroundColor Gray
        Write-Host ""
        Write-Host "       Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "    3. Close and reopen PowerShell, then re-run: .\setup.ps1" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  More info: https://chocolatey.org/install" -ForegroundColor DarkGray
        return $false
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host ""
        Write-Host "  Chocolatey needs an elevated (Admin) PowerShell to install packages." -ForegroundColor Yellow
        Write-Host "  Right-click PowerShell > 'Run as Administrator', then re-run: .\setup.ps1" -ForegroundColor Gray
        Write-Host ""
        return $false
    }

    Write-Host "  Installing $DisplayName via Chocolatey..." -ForegroundColor Cyan
    try {
        & choco install $PackageName -y 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  $DisplayName installed successfully." -ForegroundColor Green

            # Refresh PATH so we can find the newly installed exe
            $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
            $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
            $env:PATH = "$machinePath;$userPath"

            return $true
        }
    } catch {}

    Write-Host "  Chocolatey install failed." -ForegroundColor Red
    return $false
}

# ========== HELPER: MANUAL DOWNLOAD GUIDANCE ==========
function Show-ManualDownloadHelp {
    param(
        [string]$ToolName,
        [string]$Url,
        [string]$ExeName,
        [string]$InstallTip
    )

    Start-Process $Url
    Write-Host ""
    Write-Host "  Download page opened in your browser." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  After downloading:" -ForegroundColor White
    Write-Host "    - Run the installer (or extract the zip)" -ForegroundColor Gray
    Write-Host "    - $InstallTip" -ForegroundColor Gray
    Write-Host "    - Re-run .\setup.ps1 and it will find $ExeName automatically" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Alternatively, note where $ExeName ends up and choose" -ForegroundColor DarkGray
    Write-Host "  'Enter path manually' next time." -ForegroundColor DarkGray
}

# ========== DETECT / INSTALL MAKEMKV ==========
Write-Host "[1/5] Looking for MakeMKV..." -ForegroundColor White

$makemkvSearchPaths = @(
    "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe",
    "C:\Program Files\MakeMKV\makemkvcon64.exe",
    "C:\Program Files (x86)\MakeMKV\makemkvcon.exe",
    "C:\Program Files\MakeMKV\makemkvcon.exe"
)

$makemkvPath = Find-Executable -Name "makemkvcon64.exe" -SearchPaths $makemkvSearchPaths -RegistryKey "MakeMKV"
if (-not $makemkvPath) {
    $makemkvPath = Find-Executable -Name "makemkvcon.exe" -SearchPaths $makemkvSearchPaths -RegistryKey "MakeMKV"
}

if ($makemkvPath) {
    Write-Host "  Found: $makemkvPath" -ForegroundColor Green
} else {
    Write-Host "  MakeMKV not found." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  MakeMKV reads DVD and Blu-ray discs. RipDisc uses its command-line" -ForegroundColor Gray
    Write-Host "  tool (makemkvcon) to extract video files from your discs." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  How would you like to install it?" -ForegroundColor White
    Write-Host "  [1] Install via Chocolatey  - automatic, one command (recommended)" -ForegroundColor Gray
    Write-Host "  [2] Download from website   - manual install from makemkv.com" -ForegroundColor Gray
    Write-Host "  [3] Enter path manually     - if you already have it somewhere" -ForegroundColor Gray
    Write-Host "  [4] Skip for now" -ForegroundColor Gray
    $choice = Read-Host "  Choice [1-4]"

    switch ($choice) {
        "1" {
            if (Install-ViaChocolatey -PackageName "makemkv" -DisplayName "MakeMKV") {
                $makemkvPath = Find-Executable -Name "makemkvcon64.exe" -SearchPaths $makemkvSearchPaths -RegistryKey "MakeMKV"
                if (-not $makemkvPath) {
                    $makemkvPath = Find-Executable -Name "makemkvcon.exe" -SearchPaths $makemkvSearchPaths -RegistryKey "MakeMKV"
                }
            }
        }
        "2" {
            Show-ManualDownloadHelp `
                -ToolName "MakeMKV" `
                -Url "https://www.makemkv.com/download/" `
                -ExeName "makemkvcon64.exe" `
                -InstallTip "The default install location (Program Files) works fine"
        }
        "3" {
            Write-Host ""
            Write-Host "  Typical locations:" -ForegroundColor DarkGray
            Write-Host "    C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" -ForegroundColor DarkGray
            Write-Host "    C:\Program Files\MakeMKV\makemkvcon64.exe" -ForegroundColor DarkGray
            Write-Host ""
            $manual = Read-Host "  Full path to makemkvcon64.exe (or makemkvcon.exe)"
            if ($manual -and (Test-Path $manual)) {
                $makemkvPath = $manual
                Write-Host "  Using: $makemkvPath" -ForegroundColor Green
            } else {
                Write-Host "  File not found: $manual" -ForegroundColor Red
            }
        }
        default { Write-Host "  Skipped." -ForegroundColor DarkGray }
    }
}

# ========== DETECT / INSTALL HANDBRAKE CLI ==========
Write-Host "`n[2/5] Looking for HandBrakeCLI..." -ForegroundColor White

$handbrakeSearchPaths = @(
    "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe",
    "C:\Program Files\HandBrake\HandBrakeCLI.exe",
    "C:\Program Files (x86)\HandBrake\HandBrakeCLI.exe"
)

$handbrakePath = Find-Executable -Name "HandBrakeCLI.exe" -SearchPaths $handbrakeSearchPaths -RegistryKey "HandBrake"

if ($handbrakePath) {
    Write-Host "  Found: $handbrakePath" -ForegroundColor Green
} else {
    Write-Host "  HandBrakeCLI not found." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  HandBrakeCLI is the command-line version of HandBrake. RipDisc uses" -ForegroundColor Gray
    Write-Host "  it to encode the raw MKV files from MakeMKV into smaller MP4 files." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Note: This is the CLI (command-line) version, not the GUI app." -ForegroundColor White
    Write-Host "  They are separate downloads." -ForegroundColor White
    Write-Host ""
    Write-Host "  How would you like to install it?" -ForegroundColor White
    Write-Host "  [1] Install via Chocolatey  - automatic, one command (recommended)" -ForegroundColor Gray
    Write-Host "  [2] Download from website   - manual download from handbrake.fr" -ForegroundColor Gray
    Write-Host "  [3] Enter path manually     - if you already have it somewhere" -ForegroundColor Gray
    Write-Host "  [4] Skip for now" -ForegroundColor Gray
    $choice = Read-Host "  Choice [1-4]"

    switch ($choice) {
        "1" {
            if (Install-ViaChocolatey -PackageName "handbrake-cli" -DisplayName "HandBrakeCLI") {
                $handbrakePath = Find-Executable -Name "HandBrakeCLI.exe" -SearchPaths $handbrakeSearchPaths -RegistryKey "HandBrake"
            }
        }
        "2" {
            Show-ManualDownloadHelp `
                -ToolName "HandBrakeCLI" `
                -Url "https://handbrake.fr/downloads2.php" `
                -ExeName "HandBrakeCLI.exe" `
                -InstallTip "The download is a zip file - extract it and put HandBrakeCLI.exe somewhere permanent (e.g. C:\Tools\HandBrakeCLI.exe)"
        }
        "3" {
            Write-Host ""
            Write-Host "  Typical locations:" -ForegroundColor DarkGray
            Write-Host "    C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe  (if installed via Chocolatey)" -ForegroundColor DarkGray
            Write-Host "    C:\Program Files\HandBrake\HandBrakeCLI.exe     (if installed via MSI)" -ForegroundColor DarkGray
            Write-Host "    Wherever you extracted the zip download" -ForegroundColor DarkGray
            Write-Host ""
            $manual = Read-Host "  Full path to HandBrakeCLI.exe"
            if ($manual -and (Test-Path $manual)) {
                $handbrakePath = $manual
                Write-Host "  Using: $handbrakePath" -ForegroundColor Green
            } else {
                Write-Host "  File not found: $manual" -ForegroundColor Red
            }
        }
        default { Write-Host "  Skipped." -ForegroundColor DarkGray }
    }
}

# ========== CONFIGURE TEMP ROOT ==========
Write-Host "`n[3/5] Temp working directory" -ForegroundColor White
Write-Host ""
Write-Host "  MakeMKV rips raw files here before encoding. These files are large" -ForegroundColor Gray
Write-Host "  (a DVD is ~4-8 GB, a Blu-ray can be 25-50 GB) but are deleted" -ForegroundColor Gray
Write-Host "  automatically after encoding. Logs and queue files also go here." -ForegroundColor Gray
Write-Host ""

$defaultTempRoot = "C:\Video"
Write-Host "  Default: $defaultTempRoot" -ForegroundColor Gray
$tempInput = Read-Host "  Press Enter to accept, or type a new path"
$tempRoot = if ($tempInput) { $tempInput } else { $defaultTempRoot }

if (!(Test-Path $tempRoot)) {
    $create = Read-Host "  $tempRoot does not exist. Create it? [Y/n]"
    if ($create -ne "n") {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        Write-Host "  Created: $tempRoot" -ForegroundColor Green
    }
}

# ========== CONFIGURE DRIVES ==========
Write-Host "`n[4/5] Default drive letters" -ForegroundColor White
Write-Host ""
Write-Host "  Input drive  = your DVD/Blu-ray disc drive" -ForegroundColor Gray
Write-Host "  Output drive = where encoded files are saved (e.g. a media hard drive)" -ForegroundColor Gray

# Try to detect optical drives to help the user
try {
    $opticalDrives = Get-CimInstance Win32_CDROMDrive -ErrorAction SilentlyContinue
    if ($opticalDrives) {
        Write-Host ""
        Write-Host "  Detected optical drive(s):" -ForegroundColor DarkCyan
        foreach ($od in $opticalDrives) {
            $odLetter = ($od.Drive -replace '\\$', '')
            $odName = if ($od.Caption) { $od.Caption } else { "Unknown" }
            Write-Host "    $odLetter  $odName" -ForegroundColor DarkCyan
        }
    }
} catch {}

Write-Host ""
$defaultInputDrive = "D:"
Write-Host "  Default input (disc) drive: $defaultInputDrive" -ForegroundColor Gray
$inputDriveInput = Read-Host "  Press Enter to accept, or type your disc drive letter (e.g. G:)"
$inputDrive = if ($inputDriveInput) { $inputDriveInput.TrimEnd(":") + ":" } else { $defaultInputDrive }

# Show available volumes to help pick output drive
try {
    $volumes = Get-Volume -ErrorAction SilentlyContinue | Where-Object {
        $_.DriveLetter -and $_.DriveType -eq 'Fixed' -and $_.DriveLetter -ne 'C'
    } | Sort-Object DriveLetter
    if ($volumes) {
        Write-Host ""
        Write-Host "  Available storage drives:" -ForegroundColor DarkCyan
        foreach ($v in $volumes) {
            $freeGB = [math]::Round($v.SizeRemaining / 1GB, 0)
            $totalGB = [math]::Round($v.Size / 1GB, 0)
            $label = if ($v.FileSystemLabel) { $v.FileSystemLabel } else { "Local Disk" }
            Write-Host "    $($v.DriveLetter):  $label ($freeGB GB free of $totalGB GB)" -ForegroundColor DarkCyan
        }
    }
} catch {}

Write-Host ""
$defaultOutputDrive = "E:"
Write-Host "  Default output drive: $defaultOutputDrive" -ForegroundColor Gray
$outputDriveInput = Read-Host "  Press Enter to accept, or type your output drive letter (e.g. F:)"
$outputDrive = if ($outputDriveInput) { $outputDriveInput.TrimEnd(":") + ":" } else { $defaultOutputDrive }

# ========== CONFIGURE TMDB API KEY ==========
Write-Host "`n[5/5] TMDb API key (optional)" -ForegroundColor White
Write-Host ""
Write-Host "  TMDb (The Movie Database) lets RipDisc auto-detect the title," -ForegroundColor Gray
Write-Host "  type (movie/series), and season from your disc. Without it," -ForegroundColor Gray
Write-Host "  you just type the title manually with -title." -ForegroundColor Gray
Write-Host ""
Write-Host "  To get a free API key:" -ForegroundColor Gray
Write-Host "    1. Create an account at themoviedb.org" -ForegroundColor Gray
Write-Host "    2. Go to: https://www.themoviedb.org/settings/api" -ForegroundColor Gray
Write-Host "    3. Request an API key (choose 'Developer', any use case is fine)" -ForegroundColor Gray
Write-Host ""

$existingKey = $env:TMDB_API_KEY
if ($existingKey) {
    Write-Host "  Found existing TMDB_API_KEY environment variable." -ForegroundColor Green
    $tmdbKey = $existingKey
} else {
    $tmdbKey = Read-Host "  TMDb API key (or press Enter to skip)"
}

# ========== BUILD DRIVE LABELS ==========
# Detect optical drives for labelling
$driveLabels = @{}
try {
    $drives = Get-CimInstance Win32_CDROMDrive -ErrorAction SilentlyContinue
    $i = 0
    foreach ($d in $drives) {
        $letter = ($d.Drive -replace '\\$', '')
        $name = if ($d.Caption) { $d.Caption } else { "Drive $i" }
        $driveLabels["$i"] = "$letter $name"
        $i++
    }
} catch {}

if ($driveLabels.Count -eq 0) {
    $driveLabels["0"] = "Internal drive"
    $driveLabels["1"] = "External drive"
}

# ========== WRITE CONFIG ==========
$config = @{
    makemkvPath = if ($makemkvPath) { $makemkvPath } else { "" }
    handbrakePath = if ($handbrakePath) { $handbrakePath } else { "" }
    tempRoot = $tempRoot
    defaultInputDrive = $inputDrive
    defaultOutputDrive = $outputDrive
    driveLabels = $driveLabels
    tmdbApiKey = if ($tmdbKey) { $tmdbKey } else { "" }
}

$config | ConvertTo-Json -Depth 3 | Set-Content -Path $configPath -Encoding UTF8
Write-Host "`nConfig saved to: $configPath" -ForegroundColor Green

# ========== SUMMARY ==========
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Setup Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$statusColor = @{ $true = "Green"; $false = "Red" }

$makemkvOk = [bool]$makemkvPath
$handbrakeOk = [bool]$handbrakePath

Write-Host "  MakeMKV:       $(if ($makemkvOk) { $makemkvPath } else { 'NOT FOUND' })" -ForegroundColor $statusColor[$makemkvOk]
Write-Host "  HandBrakeCLI:  $(if ($handbrakeOk) { $handbrakePath } else { 'NOT FOUND' })" -ForegroundColor $statusColor[$handbrakeOk]
Write-Host "  Temp root:     $tempRoot" -ForegroundColor Green
Write-Host "  Input drive:   $inputDrive" -ForegroundColor Green
Write-Host "  Output drive:  $outputDrive" -ForegroundColor Green
Write-Host "  TMDb API key:  $(if ($tmdbKey) { 'Configured' } else { 'Not set (auto-discovery disabled)' })" -ForegroundColor $(if ($tmdbKey) { "Green" } else { "Yellow" })
Write-Host ""

if (-not $makemkvOk -or -not $handbrakeOk) {
    Write-Host "  Some tools are missing. Install them and re-run setup.ps1." -ForegroundColor Yellow
    Write-Host "  You can also edit $configPath directly." -ForegroundColor Gray
} else {
    Write-Host "  Ready to rip! Run: .\rip-disc.ps1" -ForegroundColor Green
}

Write-Host ""
