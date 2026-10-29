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
      $script:Config_MovieFolderNaming - 'title' or 'title-year' (movie / genre-film folders)
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

# Movie / genre-film folder naming (setup.ps1 step 6): 'title' = "Slow Dance",
# 'title-year' = "Slow Dance (2014)". Series folders are never changed.
$script:Config_MovieFolderNaming = if ($cfg -and "$($cfg.movieFolderNaming)" -in @('title', 'title-year')) { "$($cfg.movieFolderNaming)" } else { 'title' }

# ========== FOLDER NAMING HELPERS (shared by rip-disc.ps1 and continue-rip.ps1) ==========
# Folder under the output drive that holds a movie / genre film.
function Get-MovieParentFolder {
    param([switch]$Documentary, [switch]$Tutorial, [switch]$Fitness, [switch]$Music, [switch]$Surf, [switch]$Bluray)
    if ($Documentary) { return 'Documentaries' }
    if ($Tutorial) { return 'Tutorials' }
    if ($Fitness) { return 'Fitness' }
    if ($Music) { return 'Music' }
    if ($Surf) { return 'Surf' }
    if ($Bluray) { return 'Bluray' }
    return 'DVDs'
}

# "Slow Dance" + "2014" -> "Slow Dance (2014)". No valid 4-digit year, no change.
# Parentheses are not PowerShell wildcard characters, so -Path / -like keep working.
function Get-MediaFolderName {
    param([string]$SafeTitle, [string]$Year = "")
    if ("$Year" -match '^\d{4}$') { return "$SafeTitle ($Year)" }
    return $SafeTitle
}

# The one existing folder for this title under $ParentDir, plain or with a year, so a
# Disc 2 / extras / continue-rip run lands in the folder the first run made. $null when
# there is none, or more than one (e.g. two films of the same name) - the caller decides.
function Find-ExistingMediaFolder {
    param([string]$ParentDir, [string]$SafeTitle)
    if (-not $ParentDir -or -not (Test-Path -LiteralPath $ParentDir -PathType Container)) { return $null }
    $pattern = '^' + [regex]::Escape($SafeTitle) + '( \(\d{4}\))?$'
    $hits = @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pattern })
    if ($hits.Count -eq 1) { return $hits[0].Name }
    return $null
}

# The name THIS run gives its main file, as a regex: "-Part N" for a documentary part,
# "-Feature-BluRay" for -Bluray, else "-Feature". Version-specific on purpose: a genre
# folder can hold the DVD "-Feature" and the Blu-ray "-Feature-BluRay" side by side
# (Jellyfin shows them as two versions), so one must not count as the other's clash.
function Get-FeatureNamePattern {
    param([switch]$Bluray, [int]$Part = 0)
    if ($Part -gt 0) { return "-Part $Part\.[^.]+$" }
    if ($Bluray) { return '-Feature-BluRay\.[^.]+$' }
    return '-Feature\.[^.]+$'
}

# Any version's main file ("-Feature" or "-Feature-BluRay"). Step 3 never picks one of
# these as this disc's largest file and never moves one into extras.
$script:AnyFeaturePattern = '-Feature(-BluRay)?\.[^.]+$'

# The main-feature file already in $FolderPath (e.g. "Scarface-Feature.mp4"), or $null.
# A Disc 1 rip must never be put in a folder that already holds a film: Step 3 would see
# "Feature already exists" and sweep the new film into extras (Scarface 1932 into the
# 1983 film's folder, 2026-10-28). $Pattern is from Get-FeatureNamePattern.
function Get-ExistingFeatureFile {
    param([string]$FolderPath, [Parameter(Mandatory)][string]$Pattern)
    if (-not $FolderPath -or -not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return $null }
    $hit = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $Pattern -and $_.Extension -match '^\.(mp4|mkv|m4v|avi|mov|wmv)$' } |
        Select-Object -First 1
    if ($hit) { return $hit.Name }
    return $null
}

# Names of the image files already in $Dir (empty when it does not exist). Taken before
# a run so Step 3's image clean-up only removes images that appeared during the run and
# never an existing film's Jellyfin artwork (folder.jpg, backdrop.jpg ...) when a Disc 2,
# -Extras or continue run writes into that film's folder.
function Get-ImageFileNames {
    param([string]$Dir)
    if (-not $Dir -or -not (Test-Path -LiteralPath $Dir -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(jpg|jpeg|png|gif|bmp)$' } | ForEach-Object { $_.Name })
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
