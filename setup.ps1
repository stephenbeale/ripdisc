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
        Write-Host "Chocolatey is not installed." -ForegroundColor Yellow
        Write-Host "Install Chocolatey first: https://chocolatey.org/install" -ForegroundColor Gray
        return $false
    }

    Write-Host "Installing $DisplayName via Chocolatey..." -ForegroundColor Cyan
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "ERROR: Chocolatey install requires an elevated (Admin) PowerShell." -ForegroundColor Red
        Write-Host "Re-run setup.ps1 as Administrator, or install $DisplayName manually." -ForegroundColor Yellow
        return $false
    }

    try {
        & choco install $PackageName -y 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
        if ($LASTEXITCODE -eq 0) {
            Write-Host "$DisplayName installed successfully." -ForegroundColor Green

            # Refresh PATH so we can find the newly installed exe
            $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
            $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
            $env:PATH = "$machinePath;$userPath"

            return $true
        }
    } catch {}

    Write-Host "Chocolatey install failed." -ForegroundColor Red
    return $false
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
    Write-Host "  How would you like to install MakeMKV?" -ForegroundColor White
    Write-Host "  [1] Install via Chocolatey (choco install makemkv)" -ForegroundColor Gray
    Write-Host "  [2] Open download page (https://www.makemkv.com/download/)" -ForegroundColor Gray
    Write-Host "  [3] Enter path manually" -ForegroundColor Gray
    Write-Host "  [4] Skip (configure later)" -ForegroundColor Gray
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
            Start-Process "https://www.makemkv.com/download/"
            Write-Host "  Download page opened. After installing, re-run setup.ps1." -ForegroundColor Yellow
        }
        "3" {
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
    Write-Host "  How would you like to install HandBrakeCLI?" -ForegroundColor White
    Write-Host "  [1] Install via Chocolatey (choco install handbrake-cli)" -ForegroundColor Gray
    Write-Host "  [2] Open download page (https://handbrake.fr/downloads2.php)" -ForegroundColor Gray
    Write-Host "  [3] Enter path manually" -ForegroundColor Gray
    Write-Host "  [4] Skip (configure later)" -ForegroundColor Gray
    $choice = Read-Host "  Choice [1-4]"

    switch ($choice) {
        "1" {
            if (Install-ViaChocolatey -PackageName "handbrake-cli" -DisplayName "HandBrakeCLI") {
                $handbrakePath = Find-Executable -Name "HandBrakeCLI.exe" -SearchPaths $handbrakeSearchPaths -RegistryKey "HandBrake"
            }
        }
        "2" {
            Start-Process "https://handbrake.fr/downloads2.php"
            Write-Host "  Download page opened. After installing, re-run setup.ps1." -ForegroundColor Yellow
        }
        "3" {
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
Write-Host "`n[3/5] Temp working directory (for MakeMKV rips, logs, queue)" -ForegroundColor White

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

$defaultInputDrive = "D:"
Write-Host "  Default input (disc) drive: $defaultInputDrive" -ForegroundColor Gray
$inputDriveInput = Read-Host "  Press Enter to accept, or type your disc drive letter (e.g. G:)"
$inputDrive = if ($inputDriveInput) { $inputDriveInput.TrimEnd(":") + ":" } else { $defaultInputDrive }

$defaultOutputDrive = "E:"
Write-Host "  Default output drive: $defaultOutputDrive" -ForegroundColor Gray
$outputDriveInput = Read-Host "  Press Enter to accept, or type your output drive letter (e.g. F:)"
$outputDrive = if ($outputDriveInput) { $outputDriveInput.TrimEnd(":") + ":" } else { $defaultOutputDrive }

# ========== CONFIGURE TMDB API KEY ==========
Write-Host "`n[5/5] TMDb API key (optional - enables auto-discovery of disc titles)" -ForegroundColor White

$existingKey = $env:TMDB_API_KEY
if ($existingKey) {
    Write-Host "  Found existing TMDB_API_KEY environment variable." -ForegroundColor Green
    $tmdbKey = $existingKey
} else {
    Write-Host "  Get a free key at: https://www.themoviedb.org/settings/api" -ForegroundColor Gray
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
