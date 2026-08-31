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
    [int]$StartEpisode = 1,

    # Episode titles for a genre-series disc, in the order MakeMKV emits the files.
    # Always wins over the disc-label auto-detection below. Supply one name per episode
    # on the disc; any episodes beyond the supplied names fall back to plain numbering.
    [Parameter()]
    [string[]]$EpisodeNames = @(),

    # Skip the completion fanfare ([Console]::Beep melody) played at the end of a run.
    [Parameter()]
    [switch]$NoSound,

    # Skip ejecting the disc after the MakeMKV rip completes.
    [Parameter()]
    [switch]$NoEject,

    # Prints a clickable eBay UK sold-listings search URL for the ripped title at the
    # end of the FILE SUMMARY - not run automatically, since it's a convenience for
    # deciding what a physical disc might be worth, not part of the rip itself.
    [Parameter()]
    [switch]$CheckEbayPrice
)

# ========== LOAD CONFIG ==========
. (Join-Path $PSScriptRoot "Load-Config.ps1")
$makemkvconPath = $script:Config_MakeMkvPath

# Load System.Web for URL encoding (used to build the -CheckEbayPrice search URL)
Add-Type -AssemblyName System.Web

# Apply config defaults to parameters that weren't explicitly passed
# Captured before the defaulting below overwrites it - Stop-WithError uses this later to
# decide whether -OutputDrive belongs in a suggested continue-rip.ps1 retry command. It
# has to be captured here: $PSBoundParameters is per-function, so Stop-WithError (defined
# further down as its own function) cannot see the top-level script's copy directly.
$script:OutputDriveExplicit = $PSBoundParameters.ContainsKey('OutputDrive')
if (-not $PSBoundParameters.ContainsKey('Drive')) { $Drive = $script:Config_DefaultInputDrive }
if (-not $PSBoundParameters.ContainsKey('OutputDrive')) { $OutputDrive = $script:Config_DefaultOutputDrive }

# ========== GENRE SERIES (e.g. multi-disc documentary box sets) ==========
# -Series combined with a genre flag (-Documentary etc.) means a multi-part,
# multi-disc title that still belongs under the genre folder rather than
# Series\ - e.g. a 5-disc, 7-episode documentary boxset. This reuses the
# existing per-disc isolation and composite-file detection that -Series
# already has (see $makemkvOutputDir and the Step 2 composite check below),
# it only changes where files land and how they're named in Step 3.
# Plain -Series (fiction TV, no genre flag) is completely unaffected.
$script:GenreLabel = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } else { $null }
$script:GenreFolder = if ($Documentary) { "Documentaries" } elseif ($Tutorial) { "Tutorials" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } else { $null }
$script:IsGenreSeries = $Series -and ($null -ne $script:GenreLabel)

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
    $contentType = if ($script:IsGenreSeries) { "$($script:GenreLabel) Series" } elseif ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } elseif ($Series) { "TV Series" } elseif ($Bluray) { "Blu-ray" } else { "Movie" }
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

# Builds an eBay UK sold-listings search URL for the ripped title, so a -CheckEbayPrice
# rip can print a link the user clicks to see what copies of the physical disc have
# actually sold for. Buy It Now only, "Very Good" condition or better (LH_ItemCondition=4),
# UK sellers/location only (LH_PrefLoc=1), sold listings only (LH_Sold=1) - matches the
# exact filter combination used for the same feature in the ripaudio project.
function Get-EbaySoldListingsUrl {
    param([string]$Title, [switch]$Bluray, [switch]$Series, [int]$Season = 0)
    $formatWord = if ($Bluray) { "Blu-ray" } else { "DVD" }
    $query = "$Title $formatWord"
    if ($Series -and $Season -gt 0) { $query += " Season $Season" }
    $encodedQuery = [System.Web.HttpUtility]::UrlEncode($query)
    return "https://www.ebay.co.uk/sch/i.html?_nkw=$encodedQuery&_sacat=0&_from=R40&LH_BIN=1&LH_ItemCondition=4&LH_PrefLoc=1&rt=nc&LH_Sold=1"
}

function Show-CoffeeBadge {
    $vt = [char]0x2551
    $w  = 60
    $hz = [string]::new([char]0x2550, $w)
    $tl = [char]0x2554
    $tr = [char]0x2557
    $bl = [char]0x255A
    $br = [char]0x255D
    Write-Host ""
    Write-Host "  $tl$hz$tr" -ForegroundColor DarkGray
    Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host ("   ) ) )".PadRight($w)) -NoNewline -ForegroundColor DarkYellow; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "  (_____)  "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkYellow; Write-Host ("Enjoying this app? Consider buying me a coffee!".PadRight($w - $c.Length)) -NoNewline -ForegroundColor White; Write-Host "$vt" -ForegroundColor DarkGray
    Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host ("  |     |".PadRight($w)) -NoNewline -ForegroundColor DarkYellow; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "  |     |  "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkYellow; Write-Host (">> https://buymeacoffee.com/stephenbeale".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Yellow; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "  '-----'"; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkYellow; Write-Host ("            ^^^ click here! ^^^".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Cyan; Write-Host "$vt" -ForegroundColor DarkGray
    Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host ("".PadRight($w)) -NoNewline; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "   .----.  "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkCyan; Write-Host ("I host all my sites on SiteGround - highly".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Gray; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "   |    |  "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkCyan; Write-Host ("recommended if you want to make a site!".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Gray; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "   '----'"; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkCyan; Write-Host ("".PadRight($w - $c.Length)) -NoNewline; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "   _/  \_  "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkCyan; Write-Host (">> https://siteground.com/go/steve (affiliate)".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Yellow; Write-Host "$vt" -ForegroundColor DarkGray
    $c = "  /______\ "; Write-Host "  $vt" -NoNewline -ForegroundColor DarkGray; Write-Host $c -NoNewline -ForegroundColor DarkCyan; Write-Host ("Click to check it out and support my projects!".PadRight($w - $c.Length)) -NoNewline -ForegroundColor Cyan; Write-Host "$vt" -ForegroundColor DarkGray
    Write-Host "  $bl$hz$br" -ForegroundColor DarkGray
    Write-Host ""
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
function Write-Timestamp {
    param([string]$Label)
    $ts = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    Write-Host "[$ts] $Label" -ForegroundColor DarkGray
}

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

function Get-SafeTitle {
    # Filesystem-illegal characters in the title (e.g. a "/" in "05/06 Highlight Reel")
    # get interpreted by Windows as a path separator wherever $title is used to build a
    # path or filename - silently splitting one intended folder into two nested ones, or
    # breaking log-file creation outright ("Could not find a part of the path").
    #
    # Windows also silently drops trailing dots and trailing spaces from the final
    # component of a path when the directory/file actually gets created (e.g. "W." by
    # Oliver Stone would be created on disk as "W", not "W."). Without TrimEnd here, a
    # sanitized title would keep the dot while the real directory on disk doesn't - later
    # exact-path comparisons (e.g. the Step 3 working-directory guard) then see a
    # mismatch and fail a rip that actually succeeded. TrimEnd handles any run of
    # trailing dots/spaces in either order; the IsNullOrWhiteSpace fallback guards the
    # theoretical case of a title that is nothing but illegal characters/dots/spaces
    # once sanitized.
    #
    # Callers should build paths/filenames from this; $title itself is left untouched
    # for display text (console output, log message content, window title, TMDb
    # lookups) so the user always sees their real title, not the sanitized one.
    param([string]$Title)

    $safe = ($Title -replace '[\\/:*?"<>|]', '_').TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = $Title -replace '[\\/:*?"<>|]', '_'
    }
    return $safe
}

function Get-NormalizedDriveLetter {
    # Normalizes a -Drive/-OutputDrive value to exactly one trailing colon. A bare
    # "-match ':$'" only recognizes an already-correct "F:" and otherwise appends a
    # colon unconditionally - "F:\", "F::" or a stray trailing space (e.g. pasted from
    # elsewhere) would produce a malformed "F:\:" / "F:::" instead of being cleaned up.
    # Trim whitespace/backslash, strip any existing trailing colon(s), then append
    # exactly one.
    param([string]$DriveValue)

    return ($DriveValue.Trim().TrimEnd('\')).TrimEnd(':') + ':'
}

function Test-DriveReady {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @{ Ready = $false; Drive = "Unknown"; Message = "Cannot check drive readiness: output path is empty" }
    }

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

# Wraps text in an OSC 8 terminal hyperlink when the host understands one. Windows
# Terminal and the VS Code terminal both do; the legacy conhost window renders the
# escape sequence as visible garbage instead, so this is opt-in rather than assumed.
function Format-TerminalLink {
    param([string]$Uri, [string]$Text)

    $supportsLinks = $env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode'
    if (-not $supportsLinks) { return $Text }

    # PS 5.1 has no "`e" escape, so build ESC by code point.
    $esc = [char]27
    return "$esc]8;;$Uri$esc\$Text$esc]8;;$esc\"
}

# Closing reminder of where this session's log ended up. The literal path is always
# printed so it can be copied or pasted regardless of terminal; the clickable link is
# an extra on hosts that support it.
function Show-LogFileReminder {
    if (-not $script:LogFile) { return }

    $logPath = $script:LogFile
    Write-Host "`n--- SESSION LOG ---" -ForegroundColor Cyan

    if (Test-Path $logPath) {
        Write-Host "  $(Format-TerminalLink -Uri ([uri]$logPath).AbsoluteUri -Text $logPath)" -ForegroundColor White
        $folder = Split-Path $logPath -Parent
        Write-Host "  Folder: $(Format-TerminalLink -Uri ([uri]$folder).AbsoluteUri -Text $folder)" -ForegroundColor Gray
        Write-Host "  Open it with: notepad `"$logPath`"" -ForegroundColor DarkGray
    } else {
        # The path is chosen up front, before anything is written to it. Pointing at a
        # file that was never created is worse than saying plainly that none exists.
        Write-Host "  No log file was written this session (expected at $logPath)" -ForegroundColor DarkYellow
    }
}

# ========== DISC DISCOVERY FUNCTIONS ==========
function Get-DiscInfo {
    param([string]$DiscSource)

    try {
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

# Given the parsed MakeMKV drive list, resolves the SINGLE entry that matches both the
# disc:N index AND the target drive letter. Index alone is not guaranteed unique: a drive that
# reconnects mid-session can appear in MakeMKV's own list twice under the same index - once
# keyed by drive letter, once by raw device path (e.g. \Device\CdRom3) - seen live on a flaky
# USB DVD drive. Matching on index alone can then silently return more than one object, and
# PowerShell string interpolation joins every property from all of them with a space (e.g. a
# drive name doubled up in the "Using disc:N (...)" confirmation line).
function Select-MatchedDrive {
    param(
        [array]$DrvLines,
        [int]$MatchedIndex,
        [string]$DriveLetter
    )
    $matched = $DrvLines | Where-Object { $_.Index -eq $MatchedIndex -and $_.Letter -eq $DriveLetter } | Select-Object -First 1
    if (-not $matched) {
        # Shouldn't normally happen ($MatchedIndex is expected to have come from a Letter match
        # in the caller), but fall back to the first same-index entry rather than return nothing.
        $matched = $DrvLines | Where-Object { $_.Index -eq $MatchedIndex } | Select-Object -First 1
    }
    return $matched
}

# Polls a started Process until it exits or the timeout elapses, killing it on timeout. A plain
# `&` invocation (or a bare Start()) has no timeout mechanism of its own; this is what stops an
# external process - MakeMKV's own drive probe, specifically - from hanging the whole script
# indefinitely when a physical drive is malfunctioning. Returns $true if the process exited on
# its own within the timeout, $false if it had to be killed.
function Wait-ProcessWithTimeout {
    param(
        [System.Diagnostics.Process]$Process,
        [double]$TimeoutSec
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $Process.HasExited -and $sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        Start-Sleep -Milliseconds 250
    }
    if (-not $Process.HasExited) {
        try { $Process.Kill() } catch {}
        return $false
    }
    return $true
}

# Windows-side fallback for the disc's volume label. MakeMKV's DRV: line leaves the disc
# name empty for some drives (reproducibly so on the USB DVD units here), but Windows still
# reports the label fine. Returns "" when the drive has no letter (i.e. under -DriveIndex),
# no media, or the query fails - callers treat "" as "no name available".
function Get-DiscVolumeLabel {
    param([string]$DriveLetter)

    if (-not $DriveLetter) { return "" }
    $letter = $DriveLetter.TrimEnd(':', '\')
    if ($letter.Length -ne 1) { return "" }

    try {
        # -OperationTimeoutSec bounds the WMI call itself. Without it, a drive in a bad or
        # half-disconnected state (seen live: a USB DVD drive that had started dropping out
        # mid-rip) can make WMI's device enumeration hang for minutes waiting on THAT drive,
        # even when querying for a completely different, healthy one - this call touches every
        # CD-ROM drive on the system, not just $DriveLetter. A few seconds is generous for a
        # healthy system and turns a multi-minute hang into a quick, logged non-event.
        $vol = Get-CimInstance Win32_CDROMDrive -ErrorAction Stop -OperationTimeoutSec 5 |
            Where-Object { $_.Drive -eq "${letter}:" } |
            Select-Object -First 1
        if ($vol -and $vol.MediaLoaded -and $vol.VolumeName) { return $vol.VolumeName }
    } catch {
        # Non-fatal: episode naming is a convenience, never a reason to fail a rip.
        Write-Log "Volume label lookup failed for ${letter}: $($_.Exception.Message)"
    }
    return ""
}

# Determines what kind of disc is in a drive - Audio CD, Blu-ray, DVD-Video, or a data
# disc (CD-ROM/DVD-ROM/BD-ROM, sized by capacity as a rough hint) - independent of
# MakeMKV, so it can be shown for every drive in the listing below, not just the one
# about to be ripped. Best-effort only, like Get-DiscVolumeLabel above: never throws,
# and a failure/no-media case just returns "" so callers show nothing rather than guess.
function Get-DiscTypeLabel {
    param([string]$DriveLetter)

    if (-not $DriveLetter) { return "" }
    $letter = $DriveLetter.TrimEnd(':', '\')
    if ($letter.Length -ne 1) { return "" }

    # Audio CD has no real filesystem - Windows always reports the generic literal
    # volume label "Audio CD" for one, which Get-DiscVolumeLabel above already surfaces
    # via its own bounded WMI query, so this reuses that instead of repeating it.
    $volName = Get-DiscVolumeLabel -DriveLetter $letter
    if ($volName -eq 'Audio CD') { return "Audio CD" }

    $root = "${letter}:\"
    try {
        if (-not (Test-Path $root -ErrorAction Stop)) { return "" }
    } catch {
        return ""
    }
    if (Test-Path (Join-Path $root "BDMV") -ErrorAction SilentlyContinue) { return "Blu-ray" }
    if (Test-Path (Join-Path $root "VIDEO_TS") -ErrorAction SilentlyContinue) { return "DVD-Video" }

    # Readable filesystem, no video-disc folder structure - a data disc. Capacity is a
    # rough hint at the physical format (CD/DVD/BD-sized), not a guarantee - a
    # near-empty or multi-session disc can mislead this, so it's labelled accordingly
    # rather than stated as fact.
    try {
        $driveInfo = [System.IO.DriveInfo]::new($root)
        if ($driveInfo.IsReady) {
            $sizeGB = $driveInfo.TotalSize / 1GB
            if ($sizeGB -lt 1) { return "Data disc (CD-ROM-sized)" }
            elseif ($sizeGB -lt 10) { return "Data disc (DVD-ROM-sized)" }
            else { return "Data disc (BD-ROM-sized)" }
        }
    } catch { }
    return "Data disc"
}

# Works out the episode title for each file on a genre-series disc.
# Precedence: explicit -EpisodeNames, then the cleaned disc label (only when the disc
# yielded exactly one episode - one label cannot name several files), then nothing.
# Returns an array the same length as $FileCount; "" means "number this one only".
function Resolve-EpisodeNames {
    param(
        [int]$FileCount,
        [string[]]$Supplied = @(),
        [string]$DiscLabel = ""
    )

    # Every return uses the unary comma. Without it PowerShell unrolls a one-element
    # array into a bare string on return, and the caller's [0] would then index the
    # string and hand back its first character as the episode name.
    $names = @(for ($i = 0; $i -lt $FileCount; $i++) { "" })
    if ($FileCount -le 0) { return ,$names }

    if ($Supplied -and $Supplied.Count -gt 0) {
        for ($i = 0; $i -lt $FileCount -and $i -lt $Supplied.Count; $i++) {
            $names[$i] = "$($Supplied[$i])".Trim()
        }
        return ,$names
    }

    if ($FileCount -eq 1 -and $DiscLabel) {
        # Reuse Clean-DiscName so labels normalise the same way everywhere:
        # underscores to spaces, whitespace collapsed, title case.
        $cleaned = (Clean-DiscName -RawName $DiscLabel).CleanedTitle
        # Generic labels are worse than no name at all.
        if ($cleaned -and $cleaned -notmatch '(?i)^(dvd.?video|disc|blank|untitled)$') {
            $names[0] = $cleaned
        }
    }
    return ,$names
}

# Builds the final episode filename. Jellyfin's documented pattern is
# "Series - S01E04 - Episode Name", so the separator is " - " when a name is present.
# With no name the existing "<title>-E04.ext" shape is kept unchanged.
function Get-EpisodeFileName {
    param(
        [string]$Title,
        [string]$EpisodeTag,
        [string]$EpisodeName,
        [string]$Extension
    )

    if ($EpisodeName) {
        # Strip characters Windows will not accept in a filename.
        $safeName = ($EpisodeName -replace '[\\/:*?"<>|]', '').Trim()
        if ($safeName) { return "$Title - $EpisodeTag - $safeName$Extension" }
    }
    return "$Title-$EpisodeTag$Extension"
}

function Search-TMDb {
    param([string]$SearchTitle)

    $apiKey = if ($script:Config_TmdbApiKey) { $script:Config_TmdbApiKey } else { $env:TMDB_API_KEY }
    if (-not $apiKey) {
        Write-Host "TMDb API key not set - skipping TMDb search (run setup.ps1 or set TMDB_API_KEY)" -ForegroundColor Yellow
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
$driveLetter = Get-NormalizedDriveLetter $Drive
$driveDescription = if ($DriveIndex -ge 0) {
    $hint = if ($script:Config_DriveLabels.ContainsKey("$DriveIndex")) { $script:Config_DriveLabels["$DriveIndex"] } else { "unknown drive" }
    "Drive Index $DriveIndex ($hint)"
} else {
    "Drive $driveLetter"
}

# ========== AUTO-DISCOVERY ==========
# Build disc source string for MakeMKV
# Use disc:N format which lets MakeMKV find drives by its own enumeration index.
# WMI Win32_CDROMDrive enumeration order does NOT match MakeMKV's disc:N order,
# so we always query MakeMKV directly via `info disc:9999` to get the correct index.
# Wake up the target drive first — USB optical drives go dormant and MakeMKV stalls waiting for spin-up
$null = Test-Path "${driveLetter}\" -ErrorAction SilentlyContinue

# MakeMKV's name for the drive we are ripping from (e.g. "DVD+R-DL HL-DT-ST DVDRAM GP75N 1.02 K0GPAID0457").
# Used in Step 1 to tell our drive's read errors apart from errors MakeMKV reports for OTHER drives
# while it enumerates them at startup. Empty when -DriveIndex is used (we never see the drive list).
$script:TargetDriveName = ""

if ($DriveIndex -ge 0) {
    $discSource = "disc:$DriveIndex"
} else {
    # Find which disc:N indices are already in use by running makemkvcon processes
    $busyIndices = @()
    $mkvProcs = Get-Process -Name "makemkvcon64","makemkvcon" -ErrorAction SilentlyContinue
    foreach ($proc in $mkvProcs) {
        try {
            $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.Id)" -ErrorAction SilentlyContinue).CommandLine
            if ($cmdLine -match 'disc:(\d+)') {
                $busyIndices += [int]$Matches[1]
            }
        } catch {}
    }

    # Check for cached drive mapping (5-minute TTL) to avoid slow re-queries
    $drvCacheFile = Join-Path $env:TEMP "makemkv-drive-cache.txt"
    $drvCacheTTL = 5  # minutes
    $drvOutput = $null
    if (Test-Path $drvCacheFile) {
        $cacheAge = (Get-Date) - (Get-Item $drvCacheFile).LastWriteTime
        if ($cacheAge.TotalMinutes -lt $drvCacheTTL) {
            Write-Host "Looking up drive $driveLetter (cached)..." -ForegroundColor Gray
            $drvOutput = Get-Content $drvCacheFile
        }
    }
    if (-not $drvOutput) {
        Write-Host "Looking up drive $driveLetter in MakeMKV..." -ForegroundColor Gray
        # Run this as a real Process with a hard timeout, not a plain `&` call - `&` has no
        # timeout mechanism at all, and MakeMKV's own drive probe can hang indefinitely when a
        # physical drive is malfunctioning (seen live: a USB DVD drive that had started
        # dropping connection made this exact query hang for 5+ minutes with nothing to show
        # for it). Only stdout is redirected, matching the main rip step further down -
        # MakeMKV's error lines land on stdout, not stderr.
        $drvProc = New-Object System.Diagnostics.Process
        $drvProc.StartInfo.FileName = $makemkvconPath
        $drvProc.StartInfo.Arguments = "-r info disc:9999"
        $drvProc.StartInfo.UseShellExecute = $false
        $drvProc.StartInfo.RedirectStandardOutput = $true
        $drvProc.StartInfo.RedirectStandardError = $false
        $drvProc.StartInfo.CreateNoWindow = $true
        $drvProc.Start() | Out-Null

        # 60s, not 30s: seen live, a legitimate (not stuck) query on a USB drive that had to
        # spin up from idle took ~30-35s to complete on its own. A shorter timeout would kill a
        # query that was actually about to succeed, not just a genuinely hung one.
        $drvQueryTimeoutSec = 60
        $drvQueryExitedInTime = Wait-ProcessWithTimeout -Process $drvProc -TimeoutSec $drvQueryTimeoutSec

        if (-not $drvQueryExitedInTime) {
            # No $script:LogFile exists yet at this point - LOGGING SETUP runs later, once a
            # drive has actually been identified - so there is nothing to write this to. Make
            # the console output carry the same information a log entry would instead.
            Write-Host "`n========================================" -ForegroundColor Red
            Write-Host "MakeMKV drive query timed out after ${drvQueryTimeoutSec}s" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "No session log exists yet for this run - logging only starts once a drive is" -ForegroundColor Gray
            Write-Host "identified, and this failed before that point." -ForegroundColor Gray

            # A previous run that was Ctrl+C'd can leave its own makemkvcon64.exe still running
            # and holding the drive exclusively - PowerShell's Ctrl+C does not kill a script's
            # child processes for you. That looks identical to a hung/disconnected drive from
            # here, so surface it explicitly rather than just guessing "malfunctioning".
            $otherMkvProcs = @(Get-Process -Name "makemkvcon", "makemkvcon64" -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $drvProc.Id })
            if ($otherMkvProcs.Count -gt 0) {
                $pidList = ($otherMkvProcs | ForEach-Object { $_.Id }) -join ", "
                Write-Host "`nLikely cause: $($otherMkvProcs.Count) other MakeMKV process(es) already running (PID $pidList)." -ForegroundColor Yellow
                Write-Host "If an earlier rip was stopped with Ctrl+C, its makemkvcon process can be left" -ForegroundColor Yellow
                Write-Host "running and holding the drive - Ctrl+C does not close it for you. Close it with:" -ForegroundColor Yellow
                Write-Host "  Get-Process makemkvcon,makemkvcon64 | Stop-Process -Force" -ForegroundColor White
            } else {
                Write-Host "`nNo other MakeMKV process is currently running, so this isn't a leftover process" -ForegroundColor Gray
                Write-Host "holding the drive. Other likely causes:" -ForegroundColor Gray
                Write-Host "  - The drive is just slow to spin up from idle - simply retrying often works" -ForegroundColor Gray
                Write-Host "  - The drive/cable/hub genuinely disconnected - check the physical connection" -ForegroundColor Gray
            }

            Write-Host "`n--- RETRY OPTIONS ---" -ForegroundColor Cyan
            Write-Host "  1. Re-run the exact same command - a slow-spin-up drive often succeeds on retry" -ForegroundColor White
            Write-Host "  2. If you already know the MakeMKV drive index, skip this lookup entirely:" -ForegroundColor White
            Write-Host "       -DriveIndex <N>   (e.g. -DriveIndex 0 or -DriveIndex 1)" -ForegroundColor Gray
            Write-Host "  3. If this keeps happening on the same drive, check Task Manager for a" -ForegroundColor White
            Write-Host "     leftover makemkvcon64.exe process even after this script exits" -ForegroundColor Gray
            exit 1
        }

        # Safe to read now without blocking: the process has already exited, so its stdout
        # pipe is closed and ReadToEnd() returns immediately with whatever it wrote.
        $drvOutput = @($drvProc.StandardOutput.ReadToEnd() -split "`r?`n")
        if ($drvProc.ExitCode -ne 0) {
            Write-Host "ERROR: MakeMKV drive query failed (exit code $($drvProc.ExitCode))" -ForegroundColor Red
            exit 1
        }
        # Cache the output for subsequent runs — but only if every drive enumerated cleanly.
        # A drive that reports read errors (MSG:2003) may be missing or wrong in the DRV: list,
        # and caching that bad mapping would reuse it for the whole TTL.
        $drvHadErrors = @($drvOutput | Where-Object { "$_".Trim() -match '^MSG:2003' }).Count -gt 0
        if ($drvHadErrors) {
            Write-Host "Not caching drive list — a drive reported read errors during enumeration; will re-query next run." -ForegroundColor DarkYellow
        } else {
            $drvOutput | Set-Content $drvCacheFile -Force
        }
    }

    $matchedIndex = -1
    $drvLines = @()
    foreach ($line in $drvOutput) {
        $trimmed = "$line".Trim()
        if ($trimmed -match '^DRV:(\d+),(\d+),\d+,\d+,"([^"]*)","([^"]*)","([^"]*)"') {
            $drvIdx = [int]$Matches[1]
            $drvFlag = [int]$Matches[2]
            $drvName = $Matches[3]
            $drvDiscName = $Matches[4]
            $drvLetter = $Matches[5]
            if ($drvFlag -lt 256) {
                $isBusy = $busyIndices -contains $drvIdx
                $displayDiscName = $drvDiscName
                # For OUR target drive specifically, prefer a live Windows volume-label query
                # over the enumerated MakeMKV name. $drvOutput can be up to 5 minutes stale
                # (the cache above), so after a disc swap within that window this would
                # otherwise keep showing the PREVIOUS disc's name/label — both here in the
                # listing and later, via $script:TargetDiscLabel, in genre-series episode
                # naming. One extra per-drive-letter WMI query is cheap; re-enumerating every
                # drive on every run (what the cache exists to avoid) is not.
                # (No $DriveIndex check needed - this whole branch only runs when $DriveIndex
                # is unset; see the enclosing if/else above.)
                if ($drvLetter -eq $driveLetter) {
                    $liveDiscName = Get-DiscVolumeLabel -DriveLetter $drvLetter
                    if ($liveDiscName) { $displayDiscName = $liveDiscName }
                }
                $drvLines += [PSCustomObject]@{ Index = $drvIdx; Name = $drvName; DiscName = $displayDiscName; Letter = $drvLetter; Busy = $isBusy }
                if ($drvLetter -eq $driveLetter) {
                    $matchedIndex = $drvIdx
                }
            }
        }
    }

    # List all detected drives
    if ($drvLines.Count -gt 0) {
        Write-Host "MakeMKV drives:" -ForegroundColor Gray
        foreach ($d in $drvLines) {
            # Index alone isn't guaranteed unique: a drive that reconnects mid-session (seen
            # live, on a flaky USB DVD drive) can show up under two DRV: lines with the SAME
            # index - one keyed by drive letter, one by raw device path (e.g. \Device\CdRom3).
            # Require the letter match too, or every same-index entry gets tagged "<--".
            $marker = if ($d.Index -eq $matchedIndex -and $d.Letter -eq $driveLetter) { " <--" } else { "" }
            $busyTag = if ($d.Busy) { " (busy)" } else { "" }
            $discLabel = if ($d.DiscName) { " [$($d.DiscName)]" } else { "" }
            # Shown for every drive, not just the one about to be ripped - a quick visual
            # check to catch e.g. an audio CD sitting in the drive by mistake before a rip
            # even starts, not just after MakeMKV fails on it. Skipped for a busy drive -
            # never query one mid-rip.
            $discTypeTag = ""
            if (-not $d.Busy) {
                $discType = Get-DiscTypeLabel -DriveLetter $d.Letter
                if ($discType) { $discTypeTag = " ($discType)" }
            }
            $color = if ($d.Busy) { "DarkYellow" } else { "Gray" }
            Write-Host "  disc:$($d.Index) = $($d.Letter) - $($d.Name)$discLabel$discTypeTag$busyTag$marker" -ForegroundColor $color
        }
    }
    if ($matchedIndex -ge 0) {
        $discSource = "disc:$matchedIndex"
        # See Select-MatchedDrive: without matching on Letter too, $matchedDrv.Name below could
        # silently become a 2-element array when there's an index collision - PowerShell string
        # interpolation then joins both values with a space, e.g. "Drive Name Drive Name" in the
        # confirmation line, and the same corruption reaches $script:TargetDriveName /
        # $script:TargetDiscLabel.
        $matchedDrv = Select-MatchedDrive -DrvLines $drvLines -MatchedIndex $matchedIndex -DriveLetter $driveLetter
        # Remember MakeMKV's name for this drive so Step 1 can ignore read errors from other drives
        $script:TargetDriveName = $matchedDrv.Name
        # Remember the disc's volume label too - genre-series mode uses it to name episodes.
        # MakeMKV leaves DiscName empty for some drives (seen on USB opticals), so this may
        # be blank; Get-DiscVolumeLabel falls back to Windows for the same drive letter.
        $script:TargetDiscLabel = $matchedDrv.DiscName
        Write-Host "Using disc:$matchedIndex ($($matchedDrv.Name))" -ForegroundColor Green
    } else {
        # Clear stale cache on failure
        if (Test-Path $drvCacheFile) { Remove-Item $drvCacheFile -Force }
        Write-Host "ERROR: Drive $driveLetter not found in MakeMKV drive list." -ForegroundColor Red
        if ($drvLines.Count -gt 0) {
            Write-Host "Available MakeMKV drives:" -ForegroundColor Yellow
            foreach ($d in $drvLines) {
                Write-Host "  disc:$($d.Index) = $($d.Letter) - $($d.Name)" -ForegroundColor Gray
            }
        } else {
            Write-Host "No optical drives detected by MakeMKV." -ForegroundColor Yellow
        }
        exit 1
    }
}

if ($title -eq "") {
    # No title provided - run full discovery
    $discoverySource = $discSource
    if ($DriveIndex -ge 0) {
        $driveHint = "drive index $DriveIndex ($discoverySource)"
    } elseif ($Drive) {
        $driveHint = "$driveLetter ($discoverySource)"
    } else {
        $driveHint = "first available drive ($discoverySource)"
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "AUTO-DISCOVERY MODE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "No -title provided. Scanning $driveHint ($discoverySource)..." -ForegroundColor Yellow
    Write-Host "(This may take a minute while MakeMKV reads the disc)" -ForegroundColor Gray

    $discInfo = Get-DiscInfo -DiscSource $discoverySource

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
}
# When -title is provided, skip disc query entirely (too slow for just Blu-ray detection).
# User should pass -Bluray manually when ripping Blu-ray discs with -title.

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

# Computed here (before the confirmation prompt, not after it) so the user can see and
# abort on a bad sanitization before anything is created on disk - filesystem-illegal
# characters and Windows' silent trailing-dot/space stripping (see Get-SafeTitle) both
# mean the folder/filenames actually written can differ from $title.
$safeTitle = Get-SafeTitle $title

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Ready to rip: $title" -ForegroundColor White
if ($safeTitle -ne $title) {
    Write-Host "Folder/file name will be sanitized to: `"$safeTitle`"" -ForegroundColor Yellow
}
if ($script:IsGenreSeries) {
    if ($Season -gt 0) {
        $seasonTagPreview = "S{0:D2}" -f $Season
        Write-Host "Type: $($script:GenreLabel) Series - Season $Season ($seasonTagPreview), Disc $Disc" -ForegroundColor White
    } else {
        Write-Host "Type: $($script:GenreLabel) Series - Disc $Disc" -ForegroundColor White
    }
} elseif ($Documentary -or $Tutorial -or $Fitness -or $Music -or $Surf) {
    $genreLabel = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } else { "Surf" }
    $discTypeLabel = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: $genreLabel - $discTypeLabel$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
} elseif ($Series) {
    if ($Season -gt 0) {
        $seasonTagPreview = "S{0:D2}" -f $Season
        Write-Host "Type: TV Series - Season $Season ($seasonTagPreview), Disc $Disc" -ForegroundColor White
    } else {
        Write-Host "Type: TV Series - Disc $Disc (no season folder)" -ForegroundColor White
    }
} elseif ($Bluray) {
    $discTypeLabel = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: Blu-ray - $discTypeLabel$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
} else {
    $discTypeLabel = if ($Extras) { "Extras" } elseif ($Disc -eq 1) { "Main Feature" } else { "Special Features" }
    Write-Host "Type: Movie - $discTypeLabel$(if (-not $Extras) { " (Disc $Disc)" })" -ForegroundColor White
}
if ($script:DiscType) {
    Write-Host "Disc Format: $($script:DiscType)" -ForegroundColor Yellow
}
Write-Host "Using: $driveDescription" -ForegroundColor Yellow
Write-Host "Output Drive: $OutputDrive" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
$host.UI.RawUI.WindowTitle = "rip-disc - INPUT"
$response = Read-Host "Press Enter to continue, or Ctrl+C to abort"
Write-Timestamp "Rip started"

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
$tempRoot = $script:Config_TempRoot

# $safeTitle was already computed above, before the confirmation prompt, so the user
# gets to see and abort on a bad sanitization before anything is written to disk. Every
# path/filename built below uses it instead of $title; $title itself is left untouched
# for display text (console output, log message content, window title, TMDb lookups) so
# the user always sees their real title, not the sanitized one.

# MakeMKV temp directory - use subdirectory for multi-disc and extras rips
if ($Extras) {
    $makemkvOutputDir = "$tempRoot\$safeTitle\Extras"
} elseif ($Series -and $Season -gt 0) {
    $makemkvOutputDir = "$tempRoot\$safeTitle\Season$Season\Disc$Disc"
} else {
    $makemkvOutputDir = "$tempRoot\$safeTitle\Disc$Disc"
}

# Normalize output drive letter (add colon if missing)
$outputDriveLetter = Get-NormalizedDriveLetter $OutputDrive

# Genre types: organize into named folders (Documentaries, Tutorials, Fitness, Music)
# Genre Series: same named folder as the genre, but with per-disc isolation like Series
#               (multi-disc box sets, e.g. a documentary series spread across several discs)
# Series: organize into Season subfolders (only if Season explicitly specified)
# Movies: organize into title folder with optional extras
if ($script:IsGenreSeries) {
    $genreSeriesBaseDir = "$outputDriveLetter\$($script:GenreFolder)\$safeTitle"
    if ($Season -gt 0) {
        # Season explicitly specified - use Season subfolder
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $genreSeriesSeasonDir = Join-Path $genreSeriesBaseDir $seasonFolder
    } else {
        # No season - most documentary box sets aren't organized into seasons
        $seasonTag = $null
        $genreSeriesSeasonDir = $genreSeriesBaseDir
    }
    # Per-disc subdirectory isolates this disc's encode from other discs (same reason
    # -Series uses one). Step 3 moves the numbered episodes up and removes it afterwards.
    $finalOutputDir = Join-Path $genreSeriesSeasonDir "Disc$Disc"
} elseif ($Documentary) {
    $finalOutputDir = "$outputDriveLetter\Documentaries\$safeTitle"
} elseif ($Tutorial) {
    $finalOutputDir = "$outputDriveLetter\Tutorials\$safeTitle"
} elseif ($Fitness) {
    $finalOutputDir = "$outputDriveLetter\Fitness\$safeTitle"
} elseif ($Music) {
    $finalOutputDir = "$outputDriveLetter\Music\$safeTitle"
} elseif ($Surf) {
    $finalOutputDir = "$outputDriveLetter\Surf\$safeTitle"
} elseif ($Series) {
    $seriesBaseDir = "$outputDriveLetter\Series\$safeTitle"
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
} elseif ($Bluray) {
    $finalOutputDir = "$outputDriveLetter\Bluray\$safeTitle"
} else {
    $finalOutputDir = "$outputDriveLetter\DVDs\$safeTitle"
}

# Extras: encode directly into extras subdirectory of the title folder
if ($Extras -and -not $Series) {
    $finalOutputDir = Join-Path $finalOutputDir "extras"
}

# Fail fast if the path didn't come out usable. A malformed or empty -OutputDrive can
# make one of the Join-Path calls above emit a non-terminating error and silently
# leave $finalOutputDir null/empty (e.g. a provider-qualified-looking path such as
# "F::\..." makes Join-Path fail with "Cannot find a provider with the name 'F'" and
# return nothing, while the script carries on regardless). Left unchecked, every
# downstream Test-Path/Join-Path call on it throws a raw parameter-binding exception
# instead of a clear message. Catch it here, at construction time, instead of at each
# call site - and before spending 20+ minutes on a MakeMKV rip for nothing.
if ([string]::IsNullOrWhiteSpace($finalOutputDir) -or $finalOutputDir -notmatch '^[A-Za-z]:\\') {
    Write-Host "`nERROR: Could not build a valid output path from -OutputDrive '$OutputDrive'." -ForegroundColor Red
    Write-Host "  Resolved to: '$finalOutputDir'" -ForegroundColor Red
    Write-Host "  Expected something like 'E:\...' - check the -OutputDrive value." -ForegroundColor Yellow
    exit 1
}

# Fail fast if the destination drive itself isn't there. HandBrake encodes straight to
# this drive, so continuing here just means finding out the hard way, part-way through
# (or at the very end of) an encode that can run for many minutes - after also having
# spent time ripping the disc in Step 1.
#
# Skipped under -Queue: queue mode only rips to the local MakeMKV temp folder and
# writes an encoding job for later (see "QUEUE MODE" below) - it never touches
# $finalOutputDir in this invocation, so the destination drive is allowed to still be
# disconnected now and plugged in before -processQueue runs.
if (-not $Queue) {
    $outputDriveCheck = Test-DriveReady -Path $finalOutputDir
    if (-not $outputDriveCheck.Ready) {
        Write-Host "`nERROR: $($outputDriveCheck.Message)" -ForegroundColor Red
        exit 1
    }
}

$handbrakePath = $script:Config_HandBrakePath

# ========== LOGGING SETUP ==========
$logDir = Join-Path $script:Config_TempRoot "logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDiscLabel = if ($Extras) { "extras" } else { "disc${Disc}" }
$script:LogFile = Join-Path $logDir "${safeTitle}_${logDiscLabel}_${logTimestamp}.log"

Write-Log "========== RIP SESSION STARTED =========="
Write-Log "Title: $title"
Write-Log "Type: $(if ($script:IsGenreSeries) { "$($script:GenreLabel) Series" } elseif ($Documentary) { 'Documentary' } elseif ($Tutorial) { 'Tutorial' } elseif ($Fitness) { 'Fitness' } elseif ($Music) { 'Music' } elseif ($Surf) { 'Surf' } elseif ($Series) { 'TV Series' } elseif ($Bluray) { 'Blu-ray' } else { 'Movie' })"
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

# Builds the continue-rip.ps1 command line that would resume a failed run from its
# first incomplete step, reusing this run's own inputs, so it can be printed on
# failure for the user to copy and paste. Returns $null when there is nothing to
# resume: continue-rip.ps1 picks up AFTER the MakeMKV rip, so if Step 1 itself never
# finished there are no ripped MKV files for it to work with.
function Get-ContinueRipCommand {
    param(
        [string]$Title,
        [int[]]$RemainingStepNumbers,
        [switch]$Series,
        [int]$Season,
        [int]$Disc,
        # $null/empty means "use continue-rip.ps1's own default" - omitted from the
        # command entirely rather than printed as an explicit value.
        [string]$OutputDrive,
        [switch]$Extras,
        [switch]$Bluray,
        [switch]$Documentary,
        [switch]$Tutorial,
        [switch]$Fitness,
        [switch]$Music,
        [switch]$Surf,
        [int]$StartEpisode,
        [string[]]$EpisodeNames,
        [switch]$NoSound
    )

    # Step 1 (MakeMKV rip) has no continue-rip.ps1 equivalent - only 2/3/4 can be resumed.
    $stepToFromStep = @{ 2 = "handbrake"; 3 = "organize"; 4 = "open" }
    $firstRemaining = @($RemainingStepNumbers | Sort-Object)[0]
    if (-not $firstRemaining -or -not $stepToFromStep.ContainsKey($firstRemaining)) {
        return $null
    }

    # Escape embedded double quotes so a title/name containing one still round-trips
    # through copy-paste as a single argument.
    $quote = { param($s) '"' + ("$s" -replace '"', '`"') + '"' }

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("-title $(& $quote $Title)")
    $parts.Add("-FromStep $($stepToFromStep[$firstRemaining])")
    if ($Series) { $parts.Add("-Series") }
    if ($Season -gt 0) { $parts.Add("-Season $Season") }
    if ($Disc -ne 1) { $parts.Add("-Disc $Disc") }
    if ($OutputDrive) { $parts.Add("-OutputDrive $OutputDrive") }
    if ($Extras) { $parts.Add("-Extras") }
    if ($Bluray) { $parts.Add("-Bluray") }
    if ($Documentary) { $parts.Add("-Documentary") }
    if ($Tutorial) { $parts.Add("-Tutorial") }
    if ($Fitness) { $parts.Add("-Fitness") }
    if ($Music) { $parts.Add("-Music") }
    if ($Surf) { $parts.Add("-Surf") }
    if ($StartEpisode -ne 1) { $parts.Add("-StartEpisode $StartEpisode") }
    if ($EpisodeNames -and $EpisodeNames.Count -gt 0) {
        $quotedNames = $EpisodeNames | ForEach-Object { & $quote $_ }
        $parts.Add("-EpisodeNames $($quotedNames -join ', ')")
    }
    if ($NoSound) { $parts.Add("-NoSound") }

    return ".\continue-rip.ps1 " + ($parts -join " ")
}

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

    # Offer a ready-to-paste continue-rip.ps1 command when the MakeMKV rip itself
    # already completed - it's usually faster than the manual steps just listed.
    $stepCompleted1 = [bool]($script:CompletedSteps | Where-Object { $_.Number -eq 1 })
    if ($stepCompleted1) {
        $continueCommand = Get-ContinueRipCommand -Title $title `
            -RemainingStepNumbers @($remaining | ForEach-Object { $_.Number }) `
            -Series:$Series -Season $Season -Disc $Disc `
            -OutputDrive $(if ($script:OutputDriveExplicit) { $outputDriveLetter } else { $null }) `
            -Extras:$Extras -Bluray:$Bluray -Documentary:$Documentary -Tutorial:$Tutorial `
            -Fitness:$Fitness -Music:$Music -Surf:$Surf -StartEpisode $StartEpisode `
            -EpisodeNames $EpisodeNames -NoSound:$NoSound
        if ($continueCommand) {
            Write-Host "`n--- RETRY WITH continue-rip.ps1 ---" -ForegroundColor Cyan
            Write-Host "The MakeMKV rip already completed - resume from here instead of re-ripping the disc:" -ForegroundColor Gray
            Write-Host "  $continueCommand" -ForegroundColor White
            Write-Log "Suggested retry command: $continueCommand"
        }
    }

    # Open the relevant directory so user can see leftover files
    if ($directoryToOpen) {
        Write-Host "`n--- OPENING DIRECTORY ---" -ForegroundColor Cyan
        Write-Host "Opening: $directoryToOpen" -ForegroundColor Yellow
        Write-Host "(This is where leftover/partial files may be located)" -ForegroundColor Gray
        # Quote the path: Start-Process joins -ArgumentList on spaces, so an unquoted
        # title containing spaces is split into separate explorer.exe arguments.
        Start-Process explorer.exe -ArgumentList "`"$($directoryToOpen.TrimEnd('\'))`""
    }

    # Show recovery script path if it exists (encoding failed mid-way)
    if ($recoveryScriptPath -and (Test-Path $recoveryScriptPath)) {
        Write-Host "`n--- RECOVERY SCRIPT ---" -ForegroundColor Cyan
        Write-Host "Recovery script: $recoveryScriptPath" -ForegroundColor Yellow
        Write-Host "Run this to encode remaining files: .\$(Split-Path $recoveryScriptPath -Leaf)" -ForegroundColor White
        Write-Log "Recovery script available: $recoveryScriptPath"
    }

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "Please complete the remaining steps manually" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    # The log matters most when something has failed, so show it last rather than
    # above the error banner where it scrolls away.
    Show-LogFileReminder
    Enable-ConsoleClose
    exit 1
}

$contentType = if ($script:IsGenreSeries) { "$($script:GenreLabel) Series" } elseif ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } elseif ($Series) { "TV Series" } elseif ($Bluray) { "Blu-ray" } else { "Movie" }
# Genre types (Documentary/Tutorial/Fitness/Music/Surf) are treated like movies for file organization (Feature file, extras subfolder)
# unless combined with -Series (genre series / box sets), which are organized like Series: every file is an episode, none is "the Feature"
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
    $driveHint = if ($script:Config_DriveLabels.ContainsKey("$DriveIndex")) { $script:Config_DriveLabels["$DriveIndex"] } else { "unknown" }
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
Write-Timestamp "Step 1/4: MakeMKV rip"
Write-Host "[STEP 1/4] Starting MakeMKV rip..." -ForegroundColor Green

# Wake up dormant USB drives by accessing the drive letter (triggers spin-up)
Write-Host "Waking drive $driveLetter..." -ForegroundColor Gray
$null = Test-Path "${driveLetter}\" -ErrorAction SilentlyContinue
Write-Log "Drive wake-up: $driveLetter"

# $discSource was already set in the auto-discovery section above
if ($DriveIndex -ge 0) {
    Write-Host "Using drive index: $DriveIndex" -ForegroundColor Green
} else {
    Write-Host "Using drive: $driveLetter" -ForegroundColor Green
}

$skipMakeMkvRip = $false
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
        Write-Host "  [3] Continue with existing files (skip MakeMKV rip)" -ForegroundColor Yellow

        $choice = $null
        while ($choice -ne '1' -and $choice -ne '2' -and $choice -ne '3') {
            $choice = Read-Host "Enter 1, 2, or 3"
            if ($choice -ne '1' -and $choice -ne '2' -and $choice -ne '3') {
                Write-Host "Invalid choice. Please enter 1, 2, or 3." -ForegroundColor Red
            }
        }

        if ($choice -eq '1') {
            Write-Host "Deleting existing files..." -ForegroundColor Yellow
            $existingFiles | Remove-Item -Force
            Write-Host "Deleted $($existingFiles.Count) existing file(s)" -ForegroundColor Green
            Write-Log "User chose to delete $($existingFiles.Count) existing file(s) in $makemkvOutputDir"
        } elseif ($choice -eq '2') {
            $makemkvOutputDir = $suffixedDir
            New-Item -ItemType Directory -Path $makemkvOutputDir -Force | Out-Null
            Write-Host "Using suffixed directory: $makemkvOutputDir" -ForegroundColor Green
            Write-Log "User chose suffixed directory: $makemkvOutputDir"
        } else {
            $skipMakeMkvRip = $true
            Write-Host "Continuing with existing files — skipping MakeMKV rip" -ForegroundColor Green
            Write-Log "User chose to continue with $($existingFiles.Count) existing file(s) in $makemkvOutputDir"
        }
    } else {
        Write-Host "Directory exists (empty)" -ForegroundColor Gray
    }
} else {
    New-Item -ItemType Directory -Path $makemkvOutputDir | Out-Null
    Write-Host "Directory created successfully" -ForegroundColor Green
}

if ($skipMakeMkvRip) {
    # User chose to continue with existing files — skip MakeMKV rip and disc eject
    Write-Timestamp "Step 1/4: Skipped (using existing files)"
    Write-Host "`n[STEP 1/4] Skipped MakeMKV rip — using existing files" -ForegroundColor Cyan
    $rippedFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
    if ($null -eq $rippedFiles -or $rippedFiles.Count -eq 0) {
        Write-Host "`nERROR: No MKV files found in $makemkvOutputDir" -ForegroundColor Red
        Stop-WithError -Step "STEP 1/4: MakeMKV rip" -Message "No MKV files found in existing directory"
    }
    Write-Host "Found $($rippedFiles.Count) existing MKV file(s):" -ForegroundColor Green
    foreach ($file in $rippedFiles) {
        Write-Host "  - $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)" -ForegroundColor Gray
    }
    Write-Log "STEP 1/4: Skipped MakeMKV rip - using $($rippedFiles.Count) existing file(s) in $makemkvOutputDir"
    Complete-CurrentStep
} else {

Write-Host "`nExecuting MakeMKV command..." -ForegroundColor Yellow
Write-Host "Command: makemkvcon mkv $discSource all `"$makemkvOutputDir`" --minlength=120" -ForegroundColor Gray
Write-Log "MakeMKV command: makemkvcon mkv $discSource all `"$makemkvOutputDir`" --minlength=120"

# Stream MakeMKV output to console and capture for error analysis
# Monitors for stuck retry loops (repeated errors at the same offset) and kills the process
# if it detects MakeMKV is stuck. Titles already saved to disk are preserved.
# Stderr is not redirected — it flows to console natively. Error lines come on stdout.
$script:lastPrintedLine = $null
$stuckOffsetCount = 0
$stuckOffset = ""
$stuckThreshold = 5  # kill after this many consecutive errors at the same offset
$wasKilledForStuck = $false
# MakeMKV enumerates EVERY optical drive when it starts, so read errors from other drives arrive
# before our rip begins. Those must not trip the stuck-sector watchdog — only count offset errors
# once the rip is genuinely under way, and only when the error names the drive we are ripping.
$ripStarted = $false
# Safety net for the pre-rip phase: a disc that can never be authenticated would otherwise loop
# forever at enumeration, so kill after this many consecutive error lines with no rip in sight.
$preRipErrorCount = 0
$preRipErrorThreshold = 50
$preRipErrorDrive = ""
$wasKilledForAuth = $false

$proc = New-Object System.Diagnostics.Process
$proc.StartInfo.FileName = $makemkvconPath
$proc.StartInfo.Arguments = "mkv $discSource all `"$makemkvOutputDir`" --minlength=120"
$proc.StartInfo.UseShellExecute = $false
$proc.StartInfo.RedirectStandardOutput = $true
$proc.StartInfo.RedirectStandardError = $false
$proc.StartInfo.CreateNoWindow = $false
$proc.Start() | Out-Null

$makemkvFullOutput = New-Object System.Collections.ArrayList

# Read stdout line by line, checking for stuck retry loops
while ($null -ne ($line = $proc.StandardOutput.ReadLine())) {
    [void]$makemkvFullOutput.Add($line)

    # Detect the point where MakeMKV stops enumerating drives and starts the actual rip
    if (-not $ripStarted -and $line -match "Saving \d+ titles|Current progress|Current operation|Title #") {
        $ripStarted = $true
    }

    # Detect repeated errors at same offset (stuck retry loop)
    if ($line -match "at offset '(\d+)'") {
        $currentOffset = $Matches[1]
        # The error names the drive it came from:
        #   "... occurred while reading '<drive name>' at offset '<n>'"
        # MakeMKV scans all drives at startup, so this can be a drive we are not ripping from.
        $errorDrive = ""
        if ($line -match "occurred while reading '([^']+)' at offset '\d+'") {
            $errorDrive = $Matches[1]
        }
        # Ours unless we know our drive's name AND the error clearly names a different one.
        # With -DriveIndex we never see the drive list, so fall back to counting every error.
        $isOurDrive = ($script:TargetDriveName -eq "") -or ($errorDrive -eq "") -or ($errorDrive -eq $script:TargetDriveName)

        if ($ripStarted -and $isOurDrive) {
            if ($currentOffset -eq $stuckOffset) {
                $stuckOffsetCount++
            } else {
                $stuckOffset = $currentOffset
                $stuckOffsetCount = 1
            }
            if ($stuckOffsetCount -ge $stuckThreshold) {
                $wasKilledForStuck = $true
                Write-Host "`nWARNING: MakeMKV stuck retrying the same bad sector ($stuckOffsetCount attempts at offset $stuckOffset)" -ForegroundColor Yellow
                Write-Host "Killing MakeMKV to salvage titles already saved to disk..." -ForegroundColor Yellow
                try { $proc.Kill() } catch {}
                break
            }
        } elseif (-not $ripStarted) {
            # Still enumerating drives — never a stuck sector, but do not let it spin forever
            $preRipErrorCount++
            if ($errorDrive -ne "") { $preRipErrorDrive = $errorDrive }
            if ($preRipErrorCount -ge $preRipErrorThreshold) {
                $wasKilledForAuth = $true
                $preRipDriveLabel = if ($preRipErrorDrive -ne "") { " on '$preRipErrorDrive'" } else { "" }
                Write-Host "`nERROR: MakeMKV never got past reading the disc$preRipDriveLabel ($preRipErrorCount consecutive read errors before the rip started)" -ForegroundColor Red
                Write-Host "Killing MakeMKV — the drive could not authenticate or read the disc." -ForegroundColor Red
                Write-Log "MakeMKV killed before rip started - $preRipErrorCount consecutive read errors$preRipDriveLabel"
                try { $proc.Kill() } catch {}
                break
            }
        }
    } else {
        # Reset counters when we see a non-error line (progress is being made)
        $stuckOffsetCount = 0
        $stuckOffset = ""
        $preRipErrorCount = 0
    }

    if ($line -ne $script:lastPrintedLine) {
        Write-Host $line
        $script:lastPrintedLine = $line
    }
}

try { $proc.WaitForExit() } catch {}

# A killed process reports an unhelpful exit code, so judge it by what actually landed on disk.
# A stuck-sector kill counts as success ONLY when titles were salvaged; a kill that produced
# nothing must stay non-zero so the error analysis below runs instead of reporting a bogus success.
$salvagedFiles = @(Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue)
if ($wasKilledForStuck -and $salvagedFiles.Count -gt 0) {
    $makemkvExitCode = 0
} elseif ($wasKilledForStuck -or $wasKilledForAuth) {
    $makemkvExitCode = 1
} else {
    $makemkvExitCode = $proc.ExitCode
}

$makemkvOutputText = $makemkvFullOutput -join "`n"

# Check if MakeMKV succeeded - provide specific error messages for common issues
if ($makemkvExitCode -ne 0) {
    # Analyze output to determine the specific error
    $errorMessage = "MakeMKV exited with code $makemkvExitCode"

    # Check for the drive disconnecting (or the disc being ejected) mid-rip first - it has its
    # own unmistakable Windows error text and, unlike the checks below, means the disc WAS read
    # fine; something interrupted writing partway through, so "drive/disc not found" would be
    # actively misleading here.
    if ($makemkvOutputText -match "STATUS_DEVICE_NOT_CONNECTED" -or $makemkvOutputText -match "device (which )?does not exist") {
        $errorMessage = "The drive disconnected (or the disc was ejected) partway through the rip - check the drive's USB/power connection and that the disc is still seated, then try again"
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    }
    # Check for CSS authentication failure first — it can occur without "Failed to open disc"
    elseif ($makemkvOutputText -match "SCRAMBLED SECTOR WITHOUT AUTHENTICATION") {
        # CSS authentication failure. MakeMKV often reports only the SCSI errors and never prints
        # "Failed to open disc", so this must be checked before that branch or it never fires.
        $cssDrive = ""
        if ($makemkvOutputText -match "occurred while reading '([^']+)' at offset '\d+'") {
            $cssDrive = $Matches[1]
        }
        $cssDriveLabel = if ($cssDrive -ne "") { " Drive reporting the error: $cssDrive." } else { "" }
        $errorMessage = "Disc copy protection (CSS) - the drive could not authenticate the disc.$cssDriveLabel"
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
        Write-Host "NOTE: MakeMKV reads ALL optical drives at startup, so this may be a DIFFERENT drive from the one being ripped." -ForegroundColor Yellow
        Write-Host "Remove discs from any other optical drives and try again." -ForegroundColor Yellow
    }
    # Check for "Failed to open disc" — could be drive not found OR unreadable disc
    # If MakeMKV output mentions disc structure (IFO/BUP/VOB/VTS), the drive was found
    # but the disc itself is corrupt or unreadable
    elseif ($makemkvOutputText -match "Failed to open disc") {
        # Was previously "SCRAMBLED SECTOR WITHOUT AUTHENTICATION|ILLEGAL MODE FOR THIS
        # TRACK|Scsi error" - bare "Scsi error" is not CSS-specific at all, it's the prefix
        # MakeMKV puts on essentially every SCSI-level read failure (dirty disc, damaged
        # media, wrong disc type, drive firmware quirk, genuine protection, anything), so
        # this branch was firing - and confidently telling the user it's a CSS licence-key
        # problem - for almost any hardware-level failure. Live incident: a Blu-ray drive
        # reported "Scsi error - ILLEGAL REQUEST:ILLEGAL MODE FOR THIS TRACK" and got told
        # "Disc copy protection (CSS)" - CSS is a DVD-only scheme and doesn't even apply to
        # Blu-ray (which uses AACS/BD+), so the diagnosis couldn't have been right regardless
        # of the real cause. "SCRAMBLED SECTOR WITHOUT AUTHENTICATION" is the one signature
        # that's genuinely CSS-specific (same text the standalone branch above already keys
        # on). "ILLEGAL MODE FOR THIS TRACK" is a generic SCSI ILLEGAL REQUEST that can mean
        # copy protection (CSS on DVD, AACS/BD+ on Blu-ray) OR a dirty/damaged disc OR an
        # incompatible disc/drive combination - named honestly as "could mean any of these"
        # instead of asserted as CSS. Bare "Scsi error" is no longer matched here at all.
        if ($makemkvOutputText -match "SCRAMBLED SECTOR WITHOUT AUTHENTICATION") {
            $errorMessage = "Disc copy protection (CSS) prevented reading - the drive was found but MakeMKV could not authenticate the disc. Try reinserting the disc or check that MakeMKV has a valid licence key"
            Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
        } elseif ($makemkvOutputText -match "ILLEGAL MODE FOR THIS TRACK") {
            # This exact SCSI error is also what happens when a DVD-formatted read command
            # hits a disc that isn't DVD/BD structured at all - most commonly a plain audio
            # CD (CDDA) put in for ripping with this script by mistake instead of ripaudio's
            # rip-audio.ps1. Confirmed live: a "Batman Forever" attempt that hit this error
            # turned out to be the soundtrack CD, not the movie DVD. Reuses the same
            # Get-DiscTypeLabel helper the drive listing above uses, so there is only one
            # place that knows how to tell an audio CD apart from a video disc.
            # Only attempted when a real drive letter is known: in -DriveIndex mode,
            # $driveLetter is just $Drive's unrelated default ("D:" unless overridden) and
            # may not be the physical drive MakeMKV actually read from, so the check would
            # be unreliable there.
            $looksLikeAudioCd = ($DriveIndex -lt 0) -and ((Get-DiscTypeLabel -DriveLetter $driveLetter) -eq 'Audio CD')
            if ($looksLikeAudioCd) {
                $errorMessage = "This looks like an audio CD, not a DVD/Blu-ray - MakeMKV rips video discs, not music CDs. If this is a music CD, use the ripaudio project's rip-audio.ps1 instead (e.g. .\rip-audio.ps1 -Drive $driveLetter)."
            } else {
                $errorMessage = "The drive rejected a read command for this disc (SCSI: ILLEGAL MODE FOR THIS TRACK) - the drive was found but could not read the disc. This can mean copy protection (CSS on DVD, AACS/BD+ on Blu-ray), a dirty or damaged disc, or an incompatible disc/drive combination - not necessarily CSS specifically. Try cleaning the disc, trying a different drive, or checking that MakeMKV has a valid licence key."
            }
            Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
        } elseif ($makemkvOutputText -match "IFO|BUP|VOB|VTS") {
            $errorMessage = "Disc is corrupt or unreadable - MakeMKV found the drive but could not read the disc structure"
            Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
        } else {
            if ($DriveIndex -ge 0) {
                $errorMessage = "Drive not found: Drive index $DriveIndex does not exist or is not accessible"
            } else {
                $errorMessage = "Drive not found: $driveLetter - verify the drive letter is correct"
            }
            Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
        }
    }
    # Check for drive not found / doesn't exist (other patterns)
    elseif ($makemkvOutputText -match "no disc" -or
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
            $driveHintMsg = if ($script:Config_DriveLabels.ContainsKey("$DriveIndex")) { $script:Config_DriveLabels["$DriveIndex"] } else { "drive index $DriveIndex" }
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
        $errorMessage = "Disc is corrupt or unreadable - the disc may be damaged"
        Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    }

    Stop-WithError -Step "STEP 1/4: MakeMKV rip" -Message $errorMessage
}

$rippedFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
if ($null -eq $rippedFiles -or $rippedFiles.Count -eq 0) {
    # MakeMKV succeeded but no files created - likely no valid titles found
    $errorMessage = "No MKV files were created"

    # Check output for clues about why no files were created.
    # Device-disconnect is checked first and specifically excludes the generic "no valid title"
    # match below: MakeMKV always prints a "X titles saved, Y failed" summary at the end of a
    # rip, so "0 titles saved" appears whenever EVERY title fails to save for ANY reason - a
    # disconnected drive, a full disk, permissions, anything - not just "no disc was found".
    # Matching that bare "0 titles" substring (as this used to) misreported a drive that
    # disconnected mid-save - which MakeMKV had already read 21 titles from moments earlier -
    # as if no disc were present at all.
    if ($makemkvOutputText -match "STATUS_DEVICE_NOT_CONNECTED" -or $makemkvOutputText -match "device (which )?does not exist") {
        $errorMessage = "The drive disconnected (or the disc was ejected) while MakeMKV was saving titles - the disc itself was read fine, but nothing could be written. Check the drive's USB/power connection and try again."
    } elseif ($makemkvOutputText -match "no valid title" -or $makemkvOutputText -match "no titles found") {
        $errorMessage = "No disc detected in drive - MakeMKV could not find any valid titles"
    } elseif ($makemkvOutputText -match "copy protection" -or $makemkvOutputText -match "protected") {
        $errorMessage = "Disc may be copy-protected or encrypted - MakeMKV could not extract titles"
    } else {
        $errorMessage = "No MKV files were created - check if disc is readable and contains valid content"
    }

    Write-Host "`nERROR: $errorMessage" -ForegroundColor Red
    Stop-WithError -Step "STEP 1/4: MakeMKV rip" -Message $errorMessage
}

Write-Timestamp "Step 1/4: MakeMKV rip finished"
if ($wasKilledForStuck) {
    Write-Host "`nMakeMKV rip partially complete (killed due to stuck read error)" -ForegroundColor Yellow
    Write-Host "Files salvaged: $($rippedFiles.Count)" -ForegroundColor White
    Write-Log "MakeMKV killed after $stuckThreshold consecutive read errors at offset $stuckOffset - $($rippedFiles.Count) file(s) salvaged"
} else {
    Write-Host "`nMakeMKV rip complete!" -ForegroundColor Green
    Write-Host "Files ripped: $($rippedFiles.Count)" -ForegroundColor White
}
Write-Log "STEP 1/4: MakeMKV rip complete - $($rippedFiles.Count) file(s)"
foreach ($file in $rippedFiles) {
    Write-Host "  - $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)" -ForegroundColor Gray
    Write-Log "  Ripped: $($file.Name) ($([math]::Round($file.Length/1GB, 2)) GB)"
}
Complete-CurrentStep

# Eject disc
#
# This used to be Start-Job + Wait-Job -Timeout 15, which measured the wrong thing.
# Start-Job spawns a child powershell.exe, and when this script's own concurrent
# HandBrake encodes have the CPU pinned at 100% that spawn alone takes 18-33s
# (measured). The 15s timeout therefore expired during process startup, before the
# eject could ever be reported back - so the script announced "eject timed out,
# please eject manually" while the disc was already sitting in an open tray.
#
# Two changes fix that:
#   1. The eject is issued in-process as a direct device IOCTL. No child process,
#      no Explorer involvement, and it only touches the target drive - the old
#      Shell.Application path made the shell re-enumerate every optical drive,
#      which stalls when sibling drives are mid-rip in a concurrent session.
#   2. Success is confirmed by watching the drive go not-ready, rather than by
#      trusting a call to return within a deadline. DriveInfo.IsReady costs
#      single-digit milliseconds even under full load, so polling is cheap and,
#      unlike a wall-clock timeout, it reports what actually happened.
if ($NoEject) {
    Write-Host "`nSkipping disc eject (-NoEject)" -ForegroundColor Gray
    Write-Log "Disc eject skipped (-NoEject)"
} else {

Write-Timestamp "Ejecting disc"
Write-Host "`nEjecting disc from drive $driveLetter..." -ForegroundColor Yellow

function Test-OpticalMediaPresent {
    param([string]$Root)
    try { return (New-Object System.IO.DriveInfo $Root).IsReady } catch { return $false }
}

function Invoke-OpticalEject {
    param([string]$Root)

    if (-not ('RipDiscEject' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class RipDiscEject {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateFileW(string name, uint access, uint share,
        IntPtr security, uint disposition, uint flags, IntPtr template);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(IntPtr handle, uint code,
        IntPtr inBuffer, uint inSize, IntPtr outBuffer, uint outSize,
        out uint returned, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);

    private const uint GENERIC_READ = 0x80000000;
    private const uint FILE_SHARE_READ_WRITE = 0x00000003;
    private const uint OPEN_EXISTING = 3;
    private const uint IOCTL_STORAGE_EJECT_MEDIA = 0x002D4808;

    public static bool Eject(string driveRoot, out int error) {
        error = 0;
        string path = @"\\.\" + driveRoot.TrimEnd('\\').TrimEnd(':') + ":";
        IntPtr handle = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ_WRITE,
            IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (handle == new IntPtr(-1)) {
            error = Marshal.GetLastWin32Error();
            return false;
        }
        try {
            uint returned;
            bool ok = DeviceIoControl(handle, IOCTL_STORAGE_EJECT_MEDIA,
                IntPtr.Zero, 0, IntPtr.Zero, 0, out returned, IntPtr.Zero);
            if (!ok) { error = Marshal.GetLastWin32Error(); }
            return ok;
        } finally {
            CloseHandle(handle);
        }
    }
}
'@
    }

    $ejectError = 0
    $ok = [RipDiscEject]::Eject($Root, [ref]$ejectError)
    return @{ Ok = $ok; Error = $ejectError }
}

$ejectRoot = ($driveLetter.TrimEnd('\').TrimEnd(':')) + ":"
$ejectSuccess = $false

for ($ejectAttempt = 1; $ejectAttempt -le 2 -and -not $ejectSuccess; $ejectAttempt++) {
    if ($ejectAttempt -eq 2) {
        Write-Host "Retrying eject (attempt 2)..." -ForegroundColor Yellow
        Write-Log "Retrying disc eject for drive $driveLetter (attempt 2)"
        Start-Sleep -Seconds 2
    }

    $ejectResult = Invoke-OpticalEject -Root $ejectRoot
    if (-not $ejectResult.Ok) {
        Write-Log "Eject request for drive $driveLetter failed on attempt $ejectAttempt (win32 error $($ejectResult.Error))"
    }

    # Wait for the tray to actually open. Trays typically respond in well under a
    # second; 20s is a generous ceiling that costs nothing when the eject worked.
    for ($ejectWaited = 0; $ejectWaited -lt 20; $ejectWaited++) {
        if (-not (Test-OpticalMediaPresent $ejectRoot)) { $ejectSuccess = $true; break }
        Start-Sleep -Seconds 1
    }
}

if ($ejectSuccess) {
    Write-Host "Disc ejected successfully" -ForegroundColor Green
    Write-Log "Disc ejected from drive $driveLetter"
} else {
    Write-Host "Could not eject disc after 2 attempts - please eject manually" -ForegroundColor Yellow
    Write-Log "WARNING: Disc still present in drive $driveLetter after 2 eject attempts"
    # Show Windows dialog box so user is notified even when not watching the terminal
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Could not eject the disc for '$title' on drive $driveLetter.`n`nPlease eject it manually.",
        "RipDisc - Eject Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

} # end of -NoEject guard

} # end of MakeMKV rip + eject block


# ========== BLU-RAY MODE GUARD ==========
# A single title cannot be bigger than the disc it came from, and DVD-9 (dual
# layer) tops out at ~8.5 GB. Any MKV above that is therefore proof of a Blu-ray
# source, regardless of what was passed on the command line.
#
# This matters because without -Bluray the DVD subtitle branch runs
# (--all-subtitles --subtitle-burned=none), and Blu-ray PGS tracks get burned
# into the picture anyway despite that flag - see PR #67. Burned-in subtitles
# cannot be removed afterwards, and the source MKVs are deleted as soon as Step 2
# finishes, so this is the last point where the mistake is still cheap to fix.
#
# Checked per-file rather than on the total: MakeMKV often emits the same feature
# as more than one title, so a DVD's files can legitimately sum past 8.5 GB.
# Sits before the queue block so a queued job carries the corrected flag too.
if (-not $Bluray) {
    $dvd9CapacityBytes = 8.5GB
    $rippedMkvs = @(Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue)
    $oversizedMkvs = @($rippedMkvs | Where-Object { $_.Length -gt $dvd9CapacityBytes } | Sort-Object Length -Descending)

    if ($oversizedMkvs.Count -gt 0) {
        $biggestMkv = $oversizedMkvs[0]
        $biggestGB = [math]::Round($biggestMkv.Length / 1GB, 2)

        Write-Host "`n========================================" -ForegroundColor Yellow
        Write-Host "BLU-RAY DETECTED, BUT -Bluray WAS NOT PASSED" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "Largest ripped title: $($biggestMkv.Name) ($biggestGB GB)" -ForegroundColor White
        Write-Host "A DVD cannot hold a single title that large (DVD-9 limit is 8.5 GB)." -ForegroundColor White
        Write-Host "`nIn DVD mode HandBrake burns the PGS subtitle tracks into the picture." -ForegroundColor Red
        Write-Host "That cannot be undone once encoding has finished." -ForegroundColor Red
        Write-Log "Blu-ray guard: $($biggestMkv.Name) is $biggestGB GB (over 8.5 GB) but -Bluray was not passed"

        Write-Host "`n  [Enter] Switch to Blu-ray subtitle handling (recommended)" -ForegroundColor Green
        Write-Host "  [d]     Continue in DVD mode anyway - subtitles will be burned in" -ForegroundColor Gray
        Write-Host "  [a]     Abort and keep the ripped MKV files" -ForegroundColor Gray
        $blurayChoice = Read-Host "`nChoice"
        if ($null -ne $blurayChoice) { $blurayChoice = $blurayChoice.Trim().ToLowerInvariant() }

        if ($blurayChoice -eq 'a') {
            Write-Host "`nAborted before encoding. The ripped MKV files are kept at:" -ForegroundColor Yellow
            Write-Host "  $makemkvOutputDir" -ForegroundColor White
            Write-Host "`nResume without re-ripping the disc:" -ForegroundColor Yellow
            Write-Host "  .\continue-rip.ps1 -title `"$title`" -FromStep 2 -Bluray" -ForegroundColor Cyan
            Write-Log "Blu-ray guard: aborted before encoding - MKVs kept at $makemkvOutputDir"
            $host.UI.RawUI.WindowTitle = "$windowTitle - ABORTED"
            exit 1
        } elseif ($blurayChoice -eq 'd') {
            Write-Host "`nContinuing in DVD mode - subtitles will be burned in." -ForegroundColor Yellow
            Write-Log "Blu-ray guard: continuing in DVD mode despite Blu-ray-sized titles"
        } else {
            # Default (including an unanswered prompt) is the corrective, non-destructive choice.
            $Bluray = $true
            Write-Host "`nBlu-ray subtitle handling enabled for the rest of this run." -ForegroundColor Green
            Write-Log "Blu-ray guard: -Bluray enabled automatically for the remainder of this run"
            # $finalOutputDir was resolved before Step 1, so it is deliberately not re-routed
            # here - the folder already exists and may already hold encoded files.
            Write-Host "Output directory is unchanged: $finalOutputDir" -ForegroundColor Gray
            Write-Log "Blu-ray guard: output directory left as $finalOutputDir (resolved before the rip)"
        }
    }
}


# ========== QUEUE MODE: ADD TO QUEUE AND EXIT ==========
if ($Queue) {
    Write-Log "QUEUE MODE: Writing encoding job to queue file..."
    Write-Host "`n[QUEUE MODE] Adding encoding job to queue..." -ForegroundColor Green

    $queueFilePath = Join-Path $tempRoot "handbrake-queue.json"
    $lockFilePath = "$queueFilePath.lock"

    $entry = @{
        Title = $title
        Series = [bool]$Series
        Season = $Season
        Disc = $Disc
        Bluray = [bool]$Bluray
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

    Write-Timestamp "Job queued"
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "QUEUED!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nTitle: $title" -ForegroundColor White
    Write-Host "MKV files: $mkvCount" -ForegroundColor White
    Write-Host "Queue file: $queueFilePath" -ForegroundColor White
    Write-Host "Total jobs in queue: $($queue.Count)" -ForegroundColor White
    Write-Host "`nRun 'RipDisc -processQueue' to encode all queued jobs sequentially" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Show-CoffeeBadge

    Write-Log "QUEUE MODE: Job added to queue ($($queue.Count) total jobs)"
    Write-Log "Queue file: $queueFilePath"

    # Play triumphant fanfare to signal completion (skipped with -NoSound)
    if (-not $NoSound) {
        try {
            [Console]::Beep(523, 150)  # C5
            [Console]::Beep(659, 150)  # E5
            [Console]::Beep(784, 150)  # G5
            [Console]::Beep(1047, 300) # C6 (held)
            Start-Sleep -Milliseconds 100
            [Console]::Beep(784, 150)  # G5
            [Console]::Beep(1047, 450) # C6 (triumphant hold)
        } catch { }
    }

    Enable-ConsoleClose
    $host.UI.RawUI.WindowTitle = "$windowTitle - QUEUED"
    exit 0
}


# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
Set-CurrentStep -StepNumber 2
$script:LastWorkingDirectory = $finalOutputDir
Write-Log "STEP 2/4: Starting HandBrake encoding..."
Write-Timestamp "Step 2/4: HandBrake encoding"
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
$safeTitle = Get-SafeTitle $title
$dateStamp = Get-Date -Format "yyyy-MM-dd"
$recoveryScriptPath = Join-Path $tempRoot "recovery_${safeTitle}_${dateStamp}.ps1"
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
    $recoveryLines += "    Write-Host `"Encoding: $($mkv.Name)`" -ForegroundColor Cyan"
    if ($Bluray) {
        # Blu-ray: scan for forced/foreign-language subs only, burn them in
        $recoveryLines += "    & `$handbrakePath -i `"$inputFile`" -o `"$outputFile`" --preset `"Fast 1080p30`" --all-audio --subtitle scan --subtitle-burned --verbose=1"
    } else {
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

    Write-Timestamp "Encoding file $fileCount of $($mkvFiles.Count)"
    Write-Host "`n--- Encoding file $fileCount of $($mkvFiles.Count) ---" -ForegroundColor Cyan
    Write-Host "Input:  $($mkv.Name)" -ForegroundColor White
    Write-Host "Output: $($mkv.BaseName).mp4" -ForegroundColor White
    Write-Host "Size:   $([math]::Round($mkv.Length/1GB, 2)) GB" -ForegroundColor White
    Write-Log "Encoding file $fileCount of $($mkvFiles.Count): $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB)"

    Write-Host "`nExecuting HandBrake..." -ForegroundColor Yellow

    # Blu-ray: scan for forced/foreign-language subs and burn them in (e.g. alien dialogue in English films)
    #          Skip all other PGS subtitle tracks (image-based, get burned in despite --subtitle-burned=none)
    # DVD: include all subtitles as separate tracks (VOB subs work fine in MP4)
    if ($Bluray) {
        $handbrakeArgs = @(
            "-i", $inputFile,
            "-o", $outputFile,
            "--preset", "Fast 1080p30",
            "--all-audio",
            "--subtitle", "scan",
            "--subtitle-burned",
            "--verbose=1"
        )
    } else {
        $handbrakeArgs = @(
            "-i", $inputFile,
            "-o", $outputFile,
            "--preset", "Fast 1080p30",
            "--all-audio",
            "--all-subtitles",
            "--subtitle-burned=none",
            "--verbose=1"
        )
    }
    & $handbrakePath @handbrakeArgs
    $handbrakeExitCode = $LASTEXITCODE

    if ($handbrakeExitCode -ne 0) {
        Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "HandBrake exited with code $handbrakeExitCode while encoding $($mkv.Name)"
    }

    if (Test-Path $outputFile) {
        $encodedSize = (Get-Item $outputFile).Length
        Write-Timestamp "Encoding complete: $($mkv.Name)"
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
    $minSizeMB = 10

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

        # Clean up empty parent directories left behind (stop at temp root)
        $parentDir = Split-Path $makemkvOutputDir -Parent
        while ($parentDir -and $parentDir -ne $tempRoot -and (Test-Path $parentDir)) {
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
Write-Timestamp "Step 3/4: Organising files"
Write-Host "`n[STEP 3/4] Organizing files..." -ForegroundColor Green

# Everything below renames, moves and deletes files in the CURRENT directory. If
# this cd silently fails, that current directory is wherever the script was
# launched from - which is how a run once renamed every file in the repo folder
# it was started from, using an empty prefix. Fail loudly, and confirm we really
# landed in the target before touching anything.
try {
    Set-Location -LiteralPath $finalOutputDir -ErrorAction Stop
} catch {
    Stop-WithError -Step "STEP 3/4: Organize files" -Message "Cannot change directory to $finalOutputDir - $($_.Exception.Message)"
}
$currentPath = (Get-Location).Path.TrimEnd('\')
if ($currentPath -ne $finalOutputDir.TrimEnd('\')) {
    Stop-WithError -Step "STEP 3/4: Organize files" -Message "Refusing to organize: expected to be in $finalOutputDir but the working directory is $currentPath"
}

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

if ($script:IsGenreSeries) {
    # ========== GENRE SERIES MODE: multi-disc box set (e.g. documentary series) ==========
    # Every MKV on the disc is a distinct episode of equal standing - there is no
    # Feature/extras split. Episodes are numbered sequentially and moved up out of
    # the per-disc subdirectory into the shared title (or Season) folder, so the
    # final layout is flat and Jellyfin-friendly rather than nested per disc.
    Write-Host "`nNumbering $($script:GenreLabel.ToLower()) series episodes..." -ForegroundColor Yellow
    $seasonTag = if ($Season -gt 0) { "S{0:D2}" -f $Season } else { "" }

    # $finalOutputDir is this disc's Disc$Disc subfolder; the target is its parent
    # (the title or Season folder shared by every disc in the box set).
    $genreSeriesTargetDir = Split-Path $finalOutputDir -Parent

    # Starting episode number: an explicit -StartEpisode always wins. Otherwise, scan the
    # shared target folder for the highest existing episode number and continue after it -
    # this is what lets a box set be ripped one disc per session without the user having
    # to remember or compute where numbering left off.
    if ($PSBoundParameters.ContainsKey('StartEpisode')) {
        $nextEpisode = $StartEpisode
        Write-Host "Starting at episode $nextEpisode (explicit -StartEpisode)" -ForegroundColor Gray
        Write-Log "Genre series: starting at episode $nextEpisode (explicit -StartEpisode)"
    } else {
        # Episode names put " - " around the tag ("Title - E04 - Name.mp4"), so the
        # unseasoned pattern has to accept a dash OR a space before the E and a space,
        # a dot or end-of-name after the digits - anchoring on "-E##." would silently
        # miss every named episode and restart numbering at 1.
        $episodeNumberPattern = if ($seasonTag) { "$seasonTag`E(\d+)" } else { "(?:^|[-\s])E(\d+)(?=\s|\.|$)" }
        $existingEpisodeNumbers = @()
        if (Test-Path $genreSeriesTargetDir) {
            $existingEpisodeNumbers = Get-ChildItem -Path $genreSeriesTargetDir -File -Filter "*.mp4" |
                Where-Object { $_.Name -match $episodeNumberPattern } |
                ForEach-Object { [int]$Matches[1] }
        }
        # Measure-Object -Maximum returns a Double in Windows PowerShell 5.1 even for int
        # input - cast back to int, or the "D2" format specifier below throws at runtime.
        $nextEpisode = if ($existingEpisodeNumbers) { [int](($existingEpisodeNumbers | Measure-Object -Maximum).Maximum) + 1 } else { 1 }
        Write-Host "Starting at episode $nextEpisode (detected from existing files in $genreSeriesTargetDir)" -ForegroundColor Gray
        Write-Log "Genre series: starting at episode $nextEpisode (auto-detected from $genreSeriesTargetDir)"
    }
    $firstEpisodeThisDisc = $nextEpisode

    $episodeFiles = Get-ChildItem -File | Where-Object { $_.Extension -match '\.(mp4|mkv)$' } | Sort-Object Name

    # Episode titles: -EpisodeNames wins; otherwise fall back to the disc's own volume
    # label, which on a box set like this is the individual film's name. MakeMKV may not
    # report a label (and never does under -DriveIndex, where the drive list is skipped),
    # so ask Windows as well before giving up and numbering the file only.
    $discLabelForNames = $script:TargetDiscLabel
    if (-not $discLabelForNames -and $DriveIndex -lt 0) {
        $discLabelForNames = Get-DiscVolumeLabel -DriveLetter $driveLetter
    }
    # @() belt-and-braces: Resolve-EpisodeNames already returns an array, but a bare
    # string here would make the [index] lookup below return single characters.
    $resolvedEpisodeNames = @(Resolve-EpisodeNames -FileCount $episodeFiles.Count -Supplied $EpisodeNames -DiscLabel $discLabelForNames)
    if ($episodeFiles.Count -gt 1 -and -not ($EpisodeNames -and $EpisodeNames.Count -gt 0)) {
        Write-Host "  $($episodeFiles.Count) episodes on this disc - the disc label can only name one, so these are numbered only. Pass -EpisodeNames to title them." -ForegroundColor DarkYellow
        Write-Log "Genre series: $($episodeFiles.Count) episode files, no -EpisodeNames supplied - numbering without titles"
    }

    if (!(Test-Path $genreSeriesTargetDir)) {
        New-Item -ItemType Directory -Path $genreSeriesTargetDir -Force | Out-Null
    }

    $episodeNameIndex = 0
    foreach ($file in $episodeFiles) {
        $episodeTag = if ($seasonTag) { "$seasonTag`E{0:D2}" -f $nextEpisode } else { "E{0:D2}" -f $nextEpisode }
        $thisEpisodeName = $resolvedEpisodeNames[$episodeNameIndex]
        $episodeNameIndex++
        $candidateName = Get-EpisodeFileName -Title $title -EpisodeTag $episodeTag -EpisodeName $thisEpisodeName -Extension $file.Extension
        $uniquePath = Get-UniqueFilePath -DestDir $genreSeriesTargetDir -FileName $candidateName
        $finalName = [System.IO.Path]::GetFileName($uniquePath)
        Write-Host "  $($file.Name) -> $finalName" -ForegroundColor Gray

        $maxRetries = 10
        $retryDelay = 5
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                Move-Item -LiteralPath $file.FullName -Destination $uniquePath -ErrorAction Stop
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
        $nextEpisode++
    }
    Write-Host "Numbered episodes $firstEpisodeThisDisc-$($nextEpisode - 1), moved to $genreSeriesTargetDir" -ForegroundColor Green
    Write-Log "Numbered $($episodeFiles.Count) episode(s) ($firstEpisodeThisDisc-$($nextEpisode - 1)), moved to $genreSeriesTargetDir"

    # Clean up the now-empty per-disc subdirectory. cd to the parent first - Remove-Item
    # on the working directory itself fails with "in use" (same fix as PR #53).
    Set-Location -LiteralPath $genreSeriesTargetDir
    $remainingInDiscDir = Get-ChildItem -Path $finalOutputDir -Force -ErrorAction SilentlyContinue
    if (-not $remainingInDiscDir) {
        Remove-Item -Path $finalOutputDir -Force
        Write-Host "Removed empty disc directory: $finalOutputDir" -ForegroundColor Yellow
        Write-Log "Removed empty disc directory: $finalOutputDir"
    } else {
        Write-Host "WARNING: Disc directory not empty after move, leaving in place: $finalOutputDir" -ForegroundColor Red
        Write-Log "WARNING: Disc directory not empty after episode move: $finalOutputDir"
    }
    $finalOutputDir = $genreSeriesTargetDir
    $script:LastWorkingDirectory = $finalOutputDir
} elseif ($Series) {
    # ========== SERIES MODE: Prefix files with title + season-disc tag ==========
    # Keeps original MakeMKV filenames (t00, t01...) for episode ordering
    Write-Host "`nPrefixing series files..." -ForegroundColor Yellow
    # Read the title-folder name back from disk rather than trusting $safeTitle as text.
    # $seriesBaseDir (set above where $finalOutputDir was built) IS that folder. Matches
    # the pattern Movie mode already uses below ((Get-Item $finalOutputDir).Name) and
    # closes the whole bug class - not just the specific trailing-dot/space case
    # Get-SafeTitle now handles - since any other Windows path normalization $safeTitle
    # doesn't happen to replicate would otherwise silently desync the prefix from the
    # real on-disk name the same way "W." by Oliver Stone did.
    $dirName = (Get-Item $seriesBaseDir).Name
    $seasonTag = if ($Season -gt 0) { "S{0:D2}" -f $Season } else { "" }
    $discTag = "D$Disc"
    $prefix = "$dirName-$seasonTag-$discTag"

    $episodeFiles = Get-ChildItem -File | Where-Object {
        $_.Extension -match '\.(mp4|mkv)$'
    } | Sort-Object Name

    foreach ($file in $episodeFiles) {
        $newName = "$prefix-$($file.Name)"
        Write-Host "  $($file.Name) -> $newName" -ForegroundColor Gray
        $maxRetries = 10
        $retryDelay = 5
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
    Write-Host "Renamed $($episodeFiles.Count) file(s)" -ForegroundColor Green
    Write-Log "Prefixed $($episodeFiles.Count) file(s) with $prefix"

    # Keep files in disc subdirectory (Jellyfin scans recursively)
    Write-Host "Episodes kept in disc directory: $finalOutputDir" -ForegroundColor Green
    Write-Log "Episodes kept in disc directory: $finalOutputDir"
} else {
    # ========== MOVIE MODE: Original behavior ==========
    # prefix files with parent dir name (only if not already prefixed)
    # For disc 2+, add "Special Features-" after the movie name prefix
    if ($isMainFeatureDisc) {
        Write-Host "`nPrefixing files with directory name..." -ForegroundColor Yellow
        $dirName = (Get-Item $finalOutputDir).Name
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ($dirName + "-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
            $filesToPrefix | ForEach-Object {
                $file = $_
                if ($file.Name -like ($dirName + "_*")) {
                    # Replace underscore with hyphen (e.g. Southpaw_t01.mp4 -> Southpaw-t01.mp4)
                    $newName = $dirName + "-" + $file.Name.Substring($dirName.Length + 1)
                } else {
                    $newName = $dirName + "-" + $file.Name
                }
                $maxRetries = 10
                $retryDelay = 5
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
        # Read the title-folder name back from disk (the extras folder's own parent,
        # since $finalOutputDir here IS "<title>\extras") rather than trusting $safeTitle
        # as text - same self-correcting pattern as the Movie-mode branches above/below.
        $dirName = (Get-Item $finalOutputDir).Parent.Name
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ("$dirName-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object {
                $file = $_
                if ($file.Name -like ("$dirName" + "_*")) {
                    $newName = "$dirName-" + $file.Name.Substring($dirName.Length + 1)
                } else {
                    $newName = "$dirName-" + $file.Name
                }
                Write-Host "  - $($file.Name) -> $newName" -ForegroundColor Gray
                $maxRetries = 10
                $retryDelay = 5
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
        $dirName = (Get-Item $finalOutputDir).Name
        $filesToPrefix = Get-ChildItem -File | Where-Object { $_.Name -notlike ($dirName + "-*") }
        if ($filesToPrefix) {
            Write-Host "Files to prefix: $($filesToPrefix.Count)" -ForegroundColor White
            $filesToPrefix | ForEach-Object {
                $file = $_
                if ($file.Name -like ($dirName + "_*")) {
                    $newName = $dirName + "-Special Features-" + $file.Name.Substring($dirName.Length + 1)
                } else {
                    $newName = $dirName + "-Special Features-" + $file.Name
                }
                Write-Host "  - $($file.Name) -> $newName" -ForegroundColor Gray
                $maxRetries = 10
                $retryDelay = 5
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
                $maxRetries = 10
                $retryDelay = 5
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
Write-Timestamp "Step 4/4: Opening directory"
Write-Host "`n[STEP 4/4] Opening film directory..." -ForegroundColor Green
Write-Host "Opening: $finalOutputDir" -ForegroundColor Yellow
start $finalOutputDir
Complete-CurrentStep

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Timestamp "Complete"
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
if ($CheckEbayPrice) {
    $ebayUrl = Get-EbaySoldListingsUrl -Title $title -Bluray:$Bluray -Series:$Series -Season $Season
    Write-Host "  eBay sold prices (UK, BIN, Very Good+): $ebayUrl" -ForegroundColor White
    Write-Log "eBay sold-listings URL: $ebayUrl"
}
Write-Host "========================================" -ForegroundColor Cyan
Show-CoffeeBadge

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

# Play triumphant fanfare to signal completion (skipped with -NoSound)
if (-not $NoSound) {
    try {
        [Console]::Beep(523, 150)  # C5
        [Console]::Beep(659, 150)  # E5
        [Console]::Beep(784, 150)  # G5
        [Console]::Beep(1047, 300) # C6 (held)
        Start-Sleep -Milliseconds 100
        [Console]::Beep(784, 150)  # G5
        [Console]::Beep(1047, 450) # C6 (triumphant hold)
    } catch { }
}

Enable-ConsoleClose
$host.UI.RawUI.WindowTitle = "$windowTitle - DONE"

# Last thing on screen, after every other Write-Log has flushed, so the path is still
# visible once the rip summary has scrolled by.
Show-LogFileReminder
