<#
.SYNOPSIS
    Loads RipDisc configuration from ripdisc-config.json.
    Falls back to auto-detection if config is missing or incomplete.
    Called by rip-disc.ps1 and continue-rip.ps1 via dot-sourcing.

.DESCRIPTION
    Sets the following script-scope variables:
      $script:Config_MakeMkvPath     - Full path to makemkvcon64.exe (or makemkvcon.exe)
      $script:Config_HandBrakePath   - Full path to HandBrakeCLI.exe
      $script:Config_TempRoot        - Root temp directory (default C:\Video)
      $script:Config_DefaultInputDrive  - Default disc drive letter (default D:)
      $script:Config_DefaultOutputDrive - Default output drive letter (default E:)
      $script:Config_DriveLabels     - Hashtable of drive index -> label
      $script:Config_TmdbApiKey      - TMDb API key (may be empty)
#>

$configFilePath = Join-Path $PSScriptRoot "ripdisc-config.json"

# ========== LOAD JSON CONFIG ==========
$cfg = $null
if (Test-Path $configFilePath) {
    try {
        $cfg = Get-Content $configFilePath -Raw | ConvertFrom-Json
    } catch {
        Write-Host "WARNING: Could not parse ripdisc-config.json - using defaults" -ForegroundColor Yellow
    }
}

# ========== AUTO-DETECT HELPER ==========
function Find-ToolPath {
    param(
        [string]$ExeName,
        [string[]]$SearchPaths,
        [string]$RegistryKey
    )

    # Check PATH first
    $inPath = Get-Command $ExeName -ErrorAction SilentlyContinue
    if ($inPath) { return $inPath.Source }

    # Check registry
    if ($RegistryKey) {
        foreach ($root in @("HKLM:\SOFTWARE", "HKLM:\SOFTWARE\WOW6432Node", "HKCU:\SOFTWARE")) {
            $fullKey = Join-Path $root $RegistryKey
            $reg = Get-ItemProperty -Path $fullKey -ErrorAction SilentlyContinue
            if ($reg -and $reg.InstallLocation) {
                $candidate = Join-Path $reg.InstallLocation $ExeName
                if (Test-Path $candidate) { return $candidate }
            }
        }
    }

    # Check common paths
    foreach ($p in $SearchPaths) {
        if (Test-Path $p) { return $p }
    }

    return $null
}

# ========== RESOLVE MAKEMKV ==========
$script:Config_MakeMkvPath = $null
if ($cfg -and $cfg.makemkvPath -and (Test-Path $cfg.makemkvPath)) {
    $script:Config_MakeMkvPath = $cfg.makemkvPath
} else {
    $script:Config_MakeMkvPath = Find-ToolPath -ExeName "makemkvcon64.exe" -SearchPaths @(
        "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe",
        "C:\Program Files\MakeMKV\makemkvcon64.exe"
    ) -RegistryKey "MakeMKV"

    if (-not $script:Config_MakeMkvPath) {
        $script:Config_MakeMkvPath = Find-ToolPath -ExeName "makemkvcon.exe" -SearchPaths @(
            "C:\Program Files (x86)\MakeMKV\makemkvcon.exe",
            "C:\Program Files\MakeMKV\makemkvcon.exe"
        ) -RegistryKey "MakeMKV"
    }
}

# ========== RESOLVE HANDBRAKE ==========
$script:Config_HandBrakePath = $null
if ($cfg -and $cfg.handbrakePath -and (Test-Path $cfg.handbrakePath)) {
    $script:Config_HandBrakePath = $cfg.handbrakePath
} else {
    $script:Config_HandBrakePath = Find-ToolPath -ExeName "HandBrakeCLI.exe" -SearchPaths @(
        "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe",
        "C:\Program Files\HandBrake\HandBrakeCLI.exe",
        "C:\Program Files (x86)\HandBrake\HandBrakeCLI.exe"
    ) -RegistryKey "HandBrake"
}

# ========== RESOLVE OTHER CONFIG ==========
$script:Config_TempRoot = if ($cfg -and $cfg.tempRoot) { $cfg.tempRoot } else { "C:\Video" }
$script:Config_DefaultInputDrive = if ($cfg -and $cfg.defaultInputDrive) { $cfg.defaultInputDrive } else { "D:" }
$script:Config_DefaultOutputDrive = if ($cfg -and $cfg.defaultOutputDrive) { $cfg.defaultOutputDrive } else { "E:" }

# Drive labels
$script:Config_DriveLabels = @{}
if ($cfg -and $cfg.driveLabels) {
    $cfg.driveLabels.PSObject.Properties | ForEach-Object {
        $script:Config_DriveLabels[$_.Name] = $_.Value
    }
}
if ($script:Config_DriveLabels.Count -eq 0) {
    $script:Config_DriveLabels["0"] = "Internal drive"
    $script:Config_DriveLabels["1"] = "External drive"
}

# TMDb API key (config file takes precedence over env var)
$script:Config_TmdbApiKey = ""
if ($cfg -and $cfg.tmdbApiKey) {
    $script:Config_TmdbApiKey = $cfg.tmdbApiKey
} elseif ($env:TMDB_API_KEY) {
    $script:Config_TmdbApiKey = $env:TMDB_API_KEY
}

# ========== FIRST-RUN CHECK ==========
if (-not (Test-Path $configFilePath)) {
    Write-Host "" -ForegroundColor Yellow
    Write-Host "  No ripdisc-config.json found. Run setup.ps1 for guided setup," -ForegroundColor Yellow
    Write-Host "  or RipDisc will auto-detect tool paths." -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Yellow
}

# ========== VALIDATION ==========
if (-not $script:Config_MakeMkvPath) {
    Write-Host "`nERROR: MakeMKV (makemkvcon64.exe) not found." -ForegroundColor Red
    Write-Host "  Install MakeMKV: https://www.makemkv.com/download/" -ForegroundColor Yellow
    Write-Host "  Or run: .\setup.ps1" -ForegroundColor Yellow
    Write-Host "  Or set path in ripdisc-config.json" -ForegroundColor Yellow
    exit 1
}

if (-not $script:Config_HandBrakePath) {
    Write-Host "`nERROR: HandBrakeCLI not found." -ForegroundColor Red
    Write-Host "  Install HandBrakeCLI: https://handbrake.fr/downloads2.php" -ForegroundColor Yellow
    Write-Host "  Or run: .\setup.ps1" -ForegroundColor Yellow
    Write-Host "  Or set path in ripdisc-config.json" -ForegroundColor Yellow
    exit 1
}
