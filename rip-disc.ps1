param(
    [Parameter()]
    [string]$title = "",

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
    [switch]$Tutorial,

    [Parameter()]
    [switch]$Fitness,

    [Parameter()]
    [switch]$Music,

    [Parameter()]
    [switch]$Surf,

    [Parameter()]
    [int]$StartEpisode = 1
)

# ========== TOOL PATHS ==========
$makemkvconPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

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
    $contentType = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } elseif ($Series) { "TV Series" } else { "Movie" }
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

# ========== DISC DISCOVERY FUNCTIONS ==========
function Get-DiscInfo {
    param([string]$DiscSource)

    try {
        Write-Host "Reading disc info from $DiscSource..." -ForegroundColor Yellow
        $output = & $makemkvconPath -r info $DiscSource 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "WARNING: MakeMKV info query failed (exit code $LASTEXITCODE)" -ForegroundColor Yellow
            return $null
        }

        $discType = $null
        $discName = $null
        $volumeLabel = $null
        $titles = @()

        foreach ($line in $output) {
            # Skip ErrorRecord objects from stderr, only process string output
            if ($line -is [System.Management.Automation.ErrorRecord]) { continue }
            # Trim whitespace and \r to handle Windows line endings
            $lineStr = "$line".Trim()
            # CINFO:1 = disc type
            if ($lineStr -match '^CINFO:1,\d+,"(.+)"') {
                $discType = $Matches[1]
            }
            # CINFO:2 = disc name (best title source)
            elseif ($lineStr -match '^CINFO:2,\d+,"(.+)"') {
                $discName = $Matches[1]
            }
            # CINFO:32 = volume label (fallback)
            elseif ($lineStr -match '^CINFO:32,\d+,"(.+)"') {
                $volumeLabel = $Matches[1]
            }
            # TINFO:n,9 = duration per title
            elseif ($lineStr -match '^TINFO:(\d+),9,\d+,"(.+)"') {
                $titleIdx = [int]$Matches[1]
                while ($titles.Count -le $titleIdx) { $titles += @(@{ Duration = ""; Chapters = 0; Size = 0 }) }
                $titles[$titleIdx].Duration = $Matches[2]
            }
            # TINFO:n,8 = chapter count per title
            elseif ($lineStr -match '^TINFO:(\d+),8,\d+,"(\d+)"') {
                $titleIdx = [int]$Matches[1]
                while ($titles.Count -le $titleIdx) { $titles += @(@{ Duration = ""; Chapters = 0; Size = 0 }) }
                $titles[$titleIdx].Chapters = [int]$Matches[2]
            }
            # TINFO:n,11 = size in bytes per title
            elseif ($lineStr -match '^TINFO:(\d+),11,\d+,"(\d+)"') {
                $titleIdx = [int]$Matches[1]
                while ($titles.Count -le $titleIdx) { $titles += @(@{ Duration = ""; Chapters = 0; Size = 0 }) }
                $titles[$titleIdx].Size = [long]$Matches[2]
            }
        }

        # Use disc name if available, fall back to volume label
        if (-not $discName) { $discName = $volumeLabel }

        if (-not $discName -and -not $discType) {
            Write-Host "WARNING: Could not parse disc info from MakeMKV output" -ForegroundColor Yellow
            return $null
        }

        return @{
            DiscType = $discType
            DiscName = $discName
            VolumeLabel = $volumeLabel
            Titles = $titles
        }
    } catch {
        Write-Host "WARNING: Disc info query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Clean-DiscName {
    param([string]$RawName)

    $cleaned = $RawName

    # Extract season hint before cleaning
    $seasonHint = 0
    if ($cleaned -match '(?i)S(\d{1,2})(?:E\d|D\d|\b)') {
        $seasonHint = [int]$Matches[1]
    } elseif ($cleaned -match '(?i)Season[\s._]?(\d{1,2})') {
        $seasonHint = [int]$Matches[1]
    }

    # Extract disc hint before cleaning
    $discHint = 0
    if ($cleaned -match '(?i)D(\d{1,2})(?:\b|_)') {
        $discHint = [int]$Matches[1]
    } elseif ($cleaned -match '(?i)Disc[\s._]?(\d{1,2})') {
        $discHint = [int]$Matches[1]
    }

    # Strip known suffixes
    $cleaned = $cleaned -replace '(?i)_D\d+', ''
    $cleaned = $cleaned -replace '(?i)_WS$', ''
    $cleaned = $cleaned -replace '(?i)_FS$', ''
    $cleaned = $cleaned -replace '(?i)_SE$', ''
    $cleaned = $cleaned -replace '(?i)_CE$', ''
    $cleaned = $cleaned -replace '(?i)_DISC\d+', ''
    $cleaned = $cleaned -replace '(?i)S\d{1,2}D\d{1,2}', ''
    $cleaned = $cleaned -replace '(?i)Season[\s._]?\d+', ''
    $cleaned = $cleaned -replace '(?i)Disc[\s._]?\d+', ''

    # Replace underscores with spaces
    $cleaned = $cleaned -replace '_', ' '

    # Collapse multiple spaces and trim
    $cleaned = ($cleaned -replace '\s+', ' ').Trim()

    # Title case
    $cleaned = (Get-Culture).TextInfo.ToTitleCase($cleaned.ToLower())

    return @{
        CleanedTitle = $cleaned
        SeasonHint = $seasonHint
        DiscHint = $discHint
    }
}

function Search-TMDb {
    param([string]$SearchTitle)

    $apiKey = $env:TMDB_API_KEY
    if (-not $apiKey) {
        Write-Host "TMDB_API_KEY not set - skipping TMDb search" -ForegroundColor Yellow
        return $null
    }

    try {
        $encodedTitle = [System.Uri]::EscapeDataString($SearchTitle)
        $url = "https://api.themoviedb.org/3/search/multi?query=$encodedTitle&api_key=$apiKey"
        Write-Host "Searching TMDb for: $SearchTitle" -ForegroundColor Yellow

        $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 10
        $results = $response.results | Where-Object { $_.media_type -eq "movie" -or $_.media_type -eq "tv" }

        if (-not $results -or $results.Count -eq 0) {
            Write-Host "No TMDb results found" -ForegroundColor Yellow
            return $null
        }

        # Take top 5
        $top = @($results | Select-Object -First 5)

        if ($top.Count -eq 1) {
            $r = $top[0]
            $tmdbTitle = if ($r.media_type -eq "movie") { $r.title } else { $r.name }
            $tmdbYear = if ($r.media_type -eq "movie") { ($r.release_date -split '-')[0] } else { ($r.first_air_date -split '-')[0] }
            Write-Host "TMDb match: $tmdbTitle ($tmdbYear) [$($r.media_type)]" -ForegroundColor Green
            return @{
                Title = $tmdbTitle
                Year = $tmdbYear
                MediaType = $r.media_type
                Overview = $r.overview
            }
        }

        # Multiple results - let user pick
        Write-Host "`nTMDb results:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $top.Count; $i++) {
            $r = $top[$i]
            $tmdbTitle = if ($r.media_type -eq "movie") { $r.title } else { $r.name }
            $tmdbYear = if ($r.media_type -eq "movie") { ($r.release_date -split '-')[0] } else { ($r.first_air_date -split '-')[0] }
            $typeLabel = if ($r.media_type -eq "tv") { "TV" } else { "Movie" }
            Write-Host "  [$($i + 1)] $tmdbTitle ($tmdbYear) [$typeLabel]" -ForegroundColor White
        }
        Write-Host "  [0] None of these" -ForegroundColor Gray

        $choice = $null
        while ($null -eq $choice) {
            $input = Read-Host "Select (0-$($top.Count))"
            if ($input -match '^\d+$' -and [int]$input -ge 0 -and [int]$input -le $top.Count) {
                $choice = [int]$input
            } else {
                Write-Host "Invalid choice. Enter 0-$($top.Count)." -ForegroundColor Red
            }
        }

        if ($choice -eq 0) {
            return $null
        }

        $r = $top[$choice - 1]
        $tmdbTitle = if ($r.media_type -eq "movie") { $r.title } else { $r.name }
        $tmdbYear = if ($r.media_type -eq "movie") { ($r.release_date -split '-')[0] } else { ($r.first_air_date -split '-')[0] }
        return @{
            Title = $tmdbTitle
            Year = $tmdbYear
            MediaType = $r.media_type
            Overview = $r.overview
        }
    } catch {
        Write-Host "WARNING: TMDb search failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
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

# ========== AUTO-DISCOVERY ==========
# Build disc source string for MakeMKV queries
# When no DriveIndex specified, use disc:0 to let MakeMKV find the first available disc
# (dev:D: assumes a specific drive letter which may not be the optical drive)
if ($DriveIndex -ge 0) {
    $discSource = "disc:$DriveIndex"
    $driveHint = switch ($DriveIndex) {
        0 { "D: internal" }
        1 { "G: ASUS external" }
        default { "drive index $DriveIndex" }
    }
} else {
    $discSource = "disc:0"
    $driveHint = "first available drive"
}

if ($title -eq "") {
    # No title provided - run full discovery
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "AUTO-DISCOVERY MODE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "No -title provided. Scanning $driveHint ($discSource)..." -ForegroundColor Yellow

    $discInfo = Get-DiscInfo -DiscSource $discSource

    if ($discInfo) {
        $script:DiscType = $discInfo.DiscType
        Write-Host "`nDisc Type: $($discInfo.DiscType)" -ForegroundColor White
        Write-Host "Disc Name: $($discInfo.DiscName)" -ForegroundColor White
        if ($discInfo.VolumeLabel -and $discInfo.VolumeLabel -ne $discInfo.DiscName) {
            Write-Host "Volume Label: $($discInfo.VolumeLabel)" -ForegroundColor Gray
        }
        Write-Host "Titles on disc: $($discInfo.Titles.Count)" -ForegroundColor White

        # Auto-detect Blu-ray from disc type
        if ($discInfo.DiscType -match '(?i)blu-?ray') {
            $Bluray = $true
            Write-Host "Blu-ray detected - enabling Blu-ray mode" -ForegroundColor Green
        }

        # Clean the disc name for searching
        $cleanResult = Clean-DiscName -RawName $discInfo.DiscName
        Write-Host "`nCleaned title: $($cleanResult.CleanedTitle)" -ForegroundColor White
        if ($cleanResult.SeasonHint -gt 0) {
            Write-Host "Season hint: $($cleanResult.SeasonHint)" -ForegroundColor White
        }
        if ($cleanResult.DiscHint -gt 0) {
            Write-Host "Disc hint: $($cleanResult.DiscHint)" -ForegroundColor White
        }

        # Search TMDb if API key is available
        $tmdbResult = $null
        if ($cleanResult.CleanedTitle -and $cleanResult.CleanedTitle -ne "" -and
            $cleanResult.CleanedTitle -notmatch '(?i)^(dvd.?video|disc|blank)$') {
            $tmdbResult = Search-TMDb -SearchTitle $cleanResult.CleanedTitle
        } else {
            Write-Host "Disc name too generic for TMDb search" -ForegroundColor Yellow
        }

        # Populate metadata from discovery
        if ($tmdbResult) {
            $title = $tmdbResult.Title
            if ($tmdbResult.MediaType -eq "tv" -and -not $Series) {
                $Series = $true
                Write-Host "TV series detected - enabling Series mode" -ForegroundColor Green
            }
        } else {
            # Use cleaned disc name as title
            $title = $cleanResult.CleanedTitle
        }

        # Apply season/disc hints if not already set by user
        if ($cleanResult.SeasonHint -gt 0 -and $Season -eq 0) {
            $Season = $cleanResult.SeasonHint
        }
        if ($cleanResult.DiscHint -gt 0 -and $Disc -eq 1) {
            $Disc = $cleanResult.DiscHint
        }

        # Show discovered metadata summary
        Write-Host "`n--- Discovered Metadata ---" -ForegroundColor Cyan
        Write-Host "  Title:  $title" -ForegroundColor White
        Write-Host "  Format: $(if ($Bluray) { 'Blu-ray' } else { 'DVD' })" -ForegroundColor White
        Write-Host "  Type:   $(if ($Series) { 'TV Series' } else { 'Movie' })" -ForegroundColor White
        if ($Series) {
            if ($Season -gt 0) {
                Write-Host "  Season: $Season" -ForegroundColor White
            }
            Write-Host "  Disc:   $Disc" -ForegroundColor White
        }
        Write-Host "----------------------------" -ForegroundColor Cyan

        # Prompt for confirmation
        $discoveryChoice = $null
        while ($null -eq $discoveryChoice) {
            $input = Read-Host "[Y] Accept / [E] Edit title / [A] Abort"
            switch ($input.ToUpper()) {
                'Y' { $discoveryChoice = 'Y' }
                'E' {
                    $newTitle = Read-Host "Enter title"
                    if ($newTitle -ne "") {
                        $title = $newTitle
                    }
                    # Allow toggling Series mode
                    $seriesInput = Read-Host "Is this a TV series? (y/N)"
                    if ($seriesInput -eq 'y' -or $seriesInput -eq 'Y') {
                        $Series = $true
                        if ($Season -eq 0) {
                            $seasonInput = Read-Host "Season number (0 for none)"
                            if ($seasonInput -match '^\d+$') { $Season = [int]$seasonInput }
                        }
                    } else {
                        $Series = $false
                    }
                    $discoveryChoice = 'Y'
                }
                'A' {
                    Write-Host "Aborted." -ForegroundColor Yellow
                    exit 0
                }
                default { Write-Host "Invalid choice. Enter Y, E, or A." -ForegroundColor Red }
            }
        }
    } else {
        Write-Host "Disc info not available." -ForegroundColor Yellow
    }

    # Final fallback: manual input if still no title
    if ($title -eq "") {
        $title = Read-Host "Enter title manually"
        if ($title -eq "") {
            Write-Host "ERROR: No title provided. Cannot continue." -ForegroundColor Red
            exit 1
        }
    }
} else {
    # Title was provided - only auto-detect disc format (quick info query)
    $discInfo = Get-DiscInfo -DiscSource $discSource
    if ($discInfo) {
        $script:DiscType = $discInfo.DiscType
        if (-not $Bluray -and $discInfo.DiscType -match '(?i)blu-?ray') {
            $Bluray = $true
            Write-Host "Blu-ray detected - enabling Blu-ray mode" -ForegroundColor Green
        }
    }
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
if ($Documentary -or $Tutorial -or $Fitness -or $Music -or $Surf) {
    $genreLabel = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } else { "Surf" }
    $discType = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: $genreLabel - $discType$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
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
if ($script:DiscType) {
    Write-Host "Disc Format: $($script:DiscType)" -ForegroundColor Yellow
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

# Genre types: organize into named folders (Documentaries, Tutorials, Fitness, Music)
# Series: organize into Season subfolders (only if Season explicitly specified)
# Movies: organize into title folder with optional extras
if ($Documentary) {
    $finalOutputDir = "$outputDriveLetter\Documentaries\$title"
} elseif ($Tutorial) {
    $finalOutputDir = "$outputDriveLetter\Tutorials\$title"
} elseif ($Fitness) {
    $finalOutputDir = "$outputDriveLetter\Fitness\$title"
} elseif ($Music) {
    $finalOutputDir = "$outputDriveLetter\Music\$title"
} elseif ($Surf) {
    $finalOutputDir = "$outputDriveLetter\Surf\$title"
} elseif ($Series) {
    $seriesBaseDir = "$outputDriveLetter\Series\$title"
    if ($Season -gt 0) {
        # Season explicitly specified - use Season subfolder
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $seriesSeasonDir = Join-Path $seriesBaseDir $seasonFolder
    } else {
        # No season specified - output directly to series folder, no season tag
        $seasonTag = $null
        $seriesSeasonDir = $seriesBaseDir
    }
    # Use per-disc subdirectory to isolate concurrent disc rips (prevents rename conflicts)
    $finalOutputDir = Join-Path $seriesSeasonDir "Disc$Disc"
} else {
    $finalOutputDir = "$outputDriveLetter\DVDs\$title"
}

# Extras: encode directly into extras subdirectory of the title folder
if ($Extras -and -not $Series) {
    $finalOutputDir = Join-Path $finalOutputDir "extras"
}

$handbrakePath = "C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe"

# ========== LOGGING SETUP ==========
$logDir = "C:\Video\logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDiscLabel = if ($Extras) { "extras" } else { "disc${Disc}" }
$script:LogFile = Join-Path $logDir "${title}_${logDiscLabel}_${logTimestamp}.log"

Write-Log "========== RIP SESSION STARTED =========="
Write-Log "Title: $title"
Write-Log "Type: $(if ($Documentary) { 'Documentary' } elseif ($Tutorial) { 'Tutorial' } elseif ($Fitness) { 'Fitness' } elseif ($Music) { 'Music' } elseif ($Surf) { 'Surf' } elseif ($Series) { 'TV Series' } else { 'Movie' })"
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

$contentType = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } elseif ($Series) { "TV Series" } else { "Movie" }
# Genre types (Documentary/Tutorial/Fitness/Music/Surf) are treated like movies for file organization (Feature file, extras subfolder)
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

# $discSource was already set in the auto-discovery section above
if ($DriveIndex -ge 0) {
    Write-Host "Using drive index: $DriveIndex" -ForegroundColor Green
} else {
    Write-Host "Using disc:0 (first available optical drive)" -ForegroundColor Green
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

# Eject disc (with timeout and retry to prevent hanging if drive is busy)
Write-Host "`nEjecting disc from drive $driveLetter..." -ForegroundColor Yellow
$ejectSuccess = $false
for ($ejectAttempt = 1; $ejectAttempt -le 2; $ejectAttempt++) {
    if ($ejectAttempt -eq 2) {
        Write-Host "Retrying eject (attempt 2)..." -ForegroundColor Yellow
        Write-Log "Retrying disc eject for drive $driveLetter (attempt 2)"
        Start-Sleep -Seconds 2
    }
    $ejectJob = Start-Job -ScriptBlock {
        param($drive)
        $shell = New-Object -comObject Shell.Application
        $shell.Namespace(17).ParseName($drive).InvokeVerb("Eject")
    } -ArgumentList $driveLetter
    $ejectCompleted = $ejectJob | Wait-Job -Timeout 15
    if ($ejectCompleted) {
        Remove-Job $ejectJob -Force
        $ejectSuccess = $true
        break
    } else {
        Stop-Job $ejectJob
        Remove-Job $ejectJob -Force
    }
}
if ($ejectSuccess) {
    Write-Host "Disc ejected successfully" -ForegroundColor Green
    Write-Log "Disc ejected from drive $driveLetter"
} else {
    Write-Host "Disc eject timed out after 2 attempts - please eject manually" -ForegroundColor Yellow
    Write-Log "WARNING: Disc eject timed out for drive $driveLetter after 2 attempts"
    # Show Windows dialog box so user is notified even when not watching the terminal
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Disc eject timed out for '$title' on drive $driveLetter.`n`nIt is safe to eject the disc manually.",
        "RipDisc - Eject Timeout",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
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
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nEnjoy! Consider buying me a coffee to support continued development:" -ForegroundColor Gray
    Write-Host "https://buymeacoffee.com/stephenbeale" -ForegroundColor Cyan
    Write-Host ""

    Write-Log "QUEUE MODE: Job added to queue ($($queue.Count) total jobs)"
    Write-Log "Queue file: $queueFilePath"

    # Play triumphant fanfare to signal completion
    try {
        [Console]::Beep(523, 150)  # C5
        [Console]::Beep(659, 150)  # E5
        [Console]::Beep(784, 150)  # G5
        [Console]::Beep(1047, 300) # C6 (held)
        Start-Sleep -Milliseconds 100
        [Console]::Beep(784, 150)  # G5
        [Console]::Beep(1047, 450) # C6 (triumphant hold)
    } catch { }

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
    $others = $sortedBySize | Select-Object -Skip 1
    $sumOfOthers = ($others | Measure-Object -Property Length -Sum).Sum
    # Composite is all episodes concatenated, so its size should be close to the sum of episode files (within 70-130%)
    if ($largest.Length -ge ($sumOfOthers * 0.7) -and $largest.Length -le ($sumOfOthers * 1.3)) {
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

        # Clean up empty parent directories left behind (e.g. C:\Video\Title\, C:\Video\Title\Season1\)
        $parentDir = Split-Path $makemkvOutputDir -Parent
        while ($parentDir -and $parentDir -ne "C:\Video" -and (Test-Path $parentDir)) {
            $remaining = Get-ChildItem -Path $parentDir -Force -ErrorAction SilentlyContinue
            if ($remaining.Count -eq 0) {
                Remove-Item -Path $parentDir -Force
                Write-Host "Removed empty directory: $parentDir" -ForegroundColor Yellow
                Write-Log "Removed empty parent directory: $parentDir"
                $parentDir = Split-Path $parentDir -Parent
            } else {
                break
            }
        }
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

    # Move renamed files up from Disc subdirectory to season folder
    Write-Host "`nMoving files to season directory: $seriesSeasonDir" -ForegroundColor Yellow
    $renamedFiles = Get-ChildItem -Path $finalOutputDir -File
    foreach ($file in $renamedFiles) {
        $destPath = Join-Path $seriesSeasonDir $file.Name
        $maxRetries = 5
        $retryDelay = 3
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                Move-Item -LiteralPath $file.FullName -Destination $destPath -Force -ErrorAction Stop
                break
            } catch [System.IO.IOException] {
                if ($attempt -eq $maxRetries) {
                    Write-Host "  FAILED to move $($file.Name) after $maxRetries attempts: $_" -ForegroundColor Red
                    Write-Log "ERROR: Failed to move $($file.Name) after $maxRetries attempts: $_"
                    throw
                }
                Write-Host "  File locked: $($file.Name) - retrying in ${retryDelay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay
            }
        }
    }
    # Remove empty Disc subdirectory (cd out first — current dir is inside it)
    cd $seriesSeasonDir
    if ((Get-ChildItem -Path $finalOutputDir -Force -ErrorAction SilentlyContinue).Count -eq 0) {
        Remove-Item -Path $finalOutputDir -Force
        Write-Host "Removed empty disc directory: $finalOutputDir" -ForegroundColor Yellow
        Write-Log "Removed empty disc directory: $finalOutputDir"
    }
    # Update finalOutputDir to season folder for Step 4
    $finalOutputDir = $seriesSeasonDir
    Write-Host "Files moved to: $finalOutputDir" -ForegroundColor Green
    Write-Log "Moved files to season directory: $finalOutputDir"
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
    } elseif ($Extras) {
        # Extras disc: prefix with title only (no "-extras" or "-Special Features" in name)
        Write-Host "`nPrefixing extras files with title..." -ForegroundColor Yellow
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ("$title-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object {
                $file = $_
                $newName = "$title-" + $file.Name
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
            Write-Host "Extras prefixing complete" -ForegroundColor Green
            Write-Log "Prefixed $($filesToPrefix.Count) extras file(s) with title"
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
            foreach ($video in $nonFeatureVideos) {
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
            Write-Log "Moved $($nonFeatureVideos.Count) non-feature file(s) to extras"
        } else {
            Write-Host "No non-feature videos found" -ForegroundColor Gray
        }
    } elseif ($Extras) {
        # Extras disc: files already encoded directly into extras directory — no move needed
        Write-Host "`nExtras files already in: $finalOutputDir" -ForegroundColor Green
        Write-Log "Extras files encoded directly to extras directory — no move needed"
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
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nEnjoy! Consider buying me a coffee to support continued development:" -ForegroundColor Gray
Write-Host "https://buymeacoffee.com/stephenbeale" -ForegroundColor Cyan
Write-Host ""

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

# Play triumphant fanfare to signal completion
try {
    [Console]::Beep(523, 150)  # C5
    [Console]::Beep(659, 150)  # E5
    [Console]::Beep(784, 150)  # G5
    [Console]::Beep(1047, 300) # C6 (held)
    Start-Sleep -Milliseconds 100
    [Console]::Beep(784, 150)  # G5
    [Console]::Beep(1047, 450) # C6 (triumphant hold)
} catch { }

Enable-ConsoleClose
$host.UI.RawUI.WindowTitle = "$windowTitle - DONE"
