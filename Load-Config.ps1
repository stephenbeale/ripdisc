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
      $script:Config_LibraryFolders  - Media server library folder names on the output drive (may be empty)
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

# Movie / genre-film folder naming (setup.ps1 step 7): 'title' = "Slow Dance",
# 'title-year' = "Slow Dance (2014)". Series folders are never changed.
$script:Config_MovieFolderNaming = if ($cfg -and "$($cfg.movieFolderNaming)" -in @('title', 'title-year')) { "$($cfg.movieFolderNaming)" } else { 'title' }

# Media server library folders on the output drive (setup.ps1 step 5), e.g. DVDs,
# Documentaries, Music. Empty = not set up: the rip-time library prompt scans the drive.
$script:Config_LibraryFolders = @(if ($cfg -and $cfg.libraryFolders) { @($cfg.libraryFolders) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } })

# ========== FOLDER NAMING HELPERS (shared by rip-disc.ps1 and continue-rip.ps1) ==========
# Folder under the output drive that holds a movie / genre film.
function Get-MovieParentFolder {
    param([switch]$Documentary, [switch]$Tutorial, [switch]$Fitness, [switch]$Music, [switch]$Surf, [string]$Genre = "", [switch]$Bluray)
    if ($Documentary) { return 'Documentaries' }
    if ($Tutorial) { return 'Tutorials' }
    if ($Fitness) { return 'Fitness' }
    if ($Music) { return 'Music' }
    if ($Surf) { return 'Surf' }
    if ($Genre) { return $Genre }
    if ($Bluray) { return 'Bluray' }
    return 'DVDs'
}

# "Documentary", "Surf Series", "Basketball", "TV Series", "Blu-ray" or "Movie" - the one
# content-type label for console, log and summary lines. $GenreLabel is $script:GenreLabel
# (a built-in genre or the -Genre folder name).
function Get-ContentTypeLabel {
    param([switch]$IsGenreSeries, [string]$GenreLabel = "", [switch]$Series, [switch]$Bluray)
    if ($IsGenreSeries) { return "$GenreLabel Series" }
    if ($GenreLabel) { return $GenreLabel }
    if ($Series) { return 'TV Series' }
    if ($Bluray) { return 'Blu-ray' }
    return 'Movie'
}

# A -Genre folder name must be one plain folder: no path separators, wildcards or other
# characters Windows rejects, and not "." / "..".
function Test-GenreFolderName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Trim() -in @('.', '..')) { return $false }
    if ($Name.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) { return $false }
    return $Name -notmatch '[\[\]*?]'
}

# Is the internet reachable? Probes HTTPS on the lookup hosts with a short timeout, so an
# offline machine costs ~3 s once instead of a timeout per lookup. Tests replace this.
function Test-InternetConnection {
    param([string[]]$Hosts = @('thediscdb.com', 'v2.sg.media-imdb.com'), [int]$TimeoutMs = 3000)
    try {
        if (-not [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()) { return $false }
    } catch { }
    foreach ($h in $Hosts) {
        $client = New-Object System.Net.Sockets.TcpClient
        try {
            if ($client.ConnectAsync($h, 443).Wait($TimeoutMs) -and $client.Connected) { return $true }
        } catch {
        } finally {
            $client.Close()
        }
    }
    return $false
}

# Drives a rip can be written to (local and removable disks with a letter), for the
# output-drive prompt. Each: Letter ("F:"), Label, FreeGB. Tests pass their own list.
function Get-OutputDriveChoices {
    try {
        return @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType = 2 OR DriveType = 3' -ErrorAction Stop |
            Sort-Object DeviceID | ForEach-Object {
                [pscustomobject]@{ Letter = "$($_.DeviceID)"; Label = "$($_.VolumeName)"; FreeGB = [math]::Round(([double]$_.FreeSpace) / 1GB, 0) }
            })
    } catch {
        return @()
    }
}

# Asks which drive the rip goes to when -OutputDrive was not given. Enter (or end of
# input) takes $Default, the config's defaultOutputDrive. Returns e.g. "F:".
function Read-OutputDrive {
    param([string]$Default, [object[]]$Drives = @(), [scriptblock]$ReadInput = { param($p) Read-Host $p })
    $Default = "$Default".TrimEnd('\')
    if ($Default -notmatch ':$') { $Default = "${Default}:" }
    $Default = $Default.ToUpper()
    Write-Host "`nNo -OutputDrive given - where should the rip go?" -ForegroundColor Cyan
    foreach ($d in $Drives) {
        $mark = if ($d.Letter -ieq $Default) { "  (default)" } else { "" }
        Write-Host ("  {0} {1,-20} {2,6} GB free{3}" -f $d.Letter, $(if ($d.Label) { $d.Label } else { '(no label)' }), $d.FreeGB, $mark) -ForegroundColor White
    }
    while ($true) {
        $answer = & $ReadInput "Output drive letter (Enter = $($Default.TrimEnd(':')))"
        if ($null -eq $answer -or "$answer".Trim() -eq '') { return $Default }
        $letter = "$answer".Trim().TrimEnd('\').TrimEnd(':').ToUpper()
        if ($letter -match '^[A-Z]$') {
            $letter = "${letter}:"
            if ($Drives.Count -eq 0 -or @($Drives | Where-Object { $_.Letter -ieq $letter }).Count -gt 0) { return $letter }
        }
        Write-Host "Not an output drive: '$answer'. Type a letter from the list, or Enter for $Default." -ForegroundColor Red
    }
}

# ========== MEDIA SERVER LIBRARIES (setup.ps1 step 5, and the rip-time library prompt) ==========
# Top-level folders of the output drive that a media server (Jellyfin/Plex) uses as
# libraries: F:\DVDs, F:\Documentaries, F:\Music ... Saved by setup.ps1 as libraryFolders
# (folder names, not paths, so they follow whichever output drive is picked).

# Test-LikelyLibraryFolder / Get-DriveLibraryCandidates live in MediaLibraries.ps1 (shared with setup.ps1)
. (Join-Path $PSScriptRoot 'MediaLibraries.ps1')

# What picking library folder $Name means for the rip: the built-in genre switch for
# Documentaries/Tutorials/Fitness/Music/Surf, Series for the TV series folder, the standard
# movie folder (DVDs or Bluray, by disc type) for DVDs/Bluray, else a -Genre folder.
# Returns @{ Kind = 'movie' | 'series' | 'switch' | 'genre'; Value = switch name | folder }.
function ConvertTo-LibraryChoice {
    param([string]$Name)
    switch -Regex ($Name) {
        '^(?i)(dvds|bluray)$'      { return @{ Kind = 'movie';  Value = '' } }
        '^(?i)series$'             { return @{ Kind = 'series'; Value = 'Series' } }
        '^(?i)documentaries$'      { return @{ Kind = 'switch'; Value = 'Documentary' } }
        '^(?i)tutorials$'          { return @{ Kind = 'switch'; Value = 'Tutorial' } }
        '^(?i)fitness$'            { return @{ Kind = 'switch'; Value = 'Fitness' } }
        '^(?i)music$'              { return @{ Kind = 'switch'; Value = 'Music' } }
        '^(?i)surf$'               { return @{ Kind = 'switch'; Value = 'Surf' } }
    }
    return @{ Kind = 'genre'; Value = $Name }
}

# Asks which library the rip goes into, after the output-drive prompt. Row 1 is always
# the standard movie folder (DVDs / Bluray by disc type) and Enter takes it, so the old
# behaviour is one keypress away. $Configured = libraryFolders from setup.ps1; when empty
# the drive's folders are scanned and the likely libraries listed (marked *), with the
# others behind [O]. [N] types a new folder. Returns a ConvertTo-LibraryChoice result
# plus Folder (the folder name), or $null on [A]bort.
function Read-LibraryChoice {
    param([string]$Drive, [string[]]$Configured = @(), [object[]]$Candidates = $null,
          [scriptblock]$ReadInput = { param($p) Read-Host $p })
    $Drive = "$Drive".TrimEnd('\')
    $fromConfig = @($Configured | Where-Object { $_ }).Count -gt 0
    if ($fromConfig) {
        $all = @($Configured | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ Name = "$_"; Likely = $true } })
    } else {
        if ($null -eq $Candidates) { $Candidates = @(Get-DriveLibraryCandidates -Drive $Drive) }
        $all = @($Candidates)
    }
    # DVDs / Bluray are row 1 already
    $all = @($all | Where-Object { $_.Name -notmatch '^(?i)(dvds|bluray)$' })
    $showAll = -not $fromConfig -and @($all | Where-Object { $_.Likely }).Count -eq 0
    $show = $true
    while ($true) {
        $rows = @(if ($showAll -or $fromConfig) { $all } else { $all | Where-Object { $_.Likely } })
        $hidden = $all.Count - $rows.Count
        if ($show) {
            Write-Host "`nWhich library on ${Drive}?" -ForegroundColor Cyan
            Write-Host "  [1] Movies - $Drive\DVDs or $Drive\Bluray, by disc type  (default)" -ForegroundColor White
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $kind = ConvertTo-LibraryChoice $rows[$i].Name
                $note = if ($kind.Kind -eq 'series') { "  (TV series - S##-E## episode naming)" } elseif (-not $fromConfig -and -not $rows[$i].Likely) { "  (no video found)" } else { "" }
                Write-Host ("  [{0}] {1}\{2}{3}" -f ($i + 2), $Drive, $rows[$i].Name, $note) -ForegroundColor White
            }
            if ($hidden -gt 0) { Write-Host "  [O] Other folders on $Drive ($hidden more)" -ForegroundColor Gray }
            Write-Host "  [N] New folder" -ForegroundColor Gray
            Write-Host "  [A] Abort" -ForegroundColor Gray
            if (-not $fromConfig) { Write-Host "  (setup.ps1 can save your library folders so this list is always the same)" -ForegroundColor DarkGray }
            $show = $false
        }
        $answer = & $ReadInput "Library (Enter = 1)"
        if ($null -eq $answer -or "$answer".Trim() -eq '' -or "$answer".Trim() -eq '1') { return @{ Kind = 'movie'; Value = ''; Folder = '' } }
        $answer = "$answer".Trim()
        if ($answer -match '^\d+$' -and [int]$answer -ge 2 -and [int]$answer -le $rows.Count + 1) {
            $name = $rows[[int]$answer - 2].Name
            $c = ConvertTo-LibraryChoice $name
            $c.Folder = $name
            return $c
        }
        $key = $answer.ToUpper()
        if ($key -eq 'O' -and $hidden -gt 0) { $showAll = $true; $show = $true; continue }
        if ($key -eq 'A') { return $null }
        if ($key -eq 'N') {
            $text = & $ReadInput "New folder name under $Drive\ (blank to go back)"
            if ($null -eq $text) { return @{ Kind = 'movie'; Value = ''; Folder = '' } }
            $text = "$text".Trim()
            if (-not $text) { $show = $true; continue }
            if (Test-GenreFolderName $text) { $c = ConvertTo-LibraryChoice $text; $c.Folder = $text; return $c }
            Write-Host "Not a usable folder name: '$text'" -ForegroundColor Red
            continue
        }
        Write-Host "Invalid choice. Enter 1-$($rows.Count + 1)$(if ($hidden -gt 0) { ', O' }), N or A." -ForegroundColor Red
    }
}

# Spelling- and word-tolerant "is this search hit the same title?" for the year prompt:
# the same title (Test-NearMissTitle), one title's words all inside the other's
# ("Scarface" / "Scarface: The World Is Yours"), or most words shared. Real case
# 2026-11-01: "Duxford the Second World War Years" got "Gundu" and "Apocalypse: The
# Second World War" as year choices - none of them that film.
function Test-SimilarTitle {
    param([string]$A, [string]$B)
    if (Test-NearMissTitle $A $B) { return $true }
    $stop = @('the', 'a', 'an', 'of', 'and', 'in', 'on', 'at', 'to', 'for', 'with')
    $wa = @((("$A".ToLowerInvariant() -replace '&', ' and ') -replace '[^\p{L}\p{Nd}]+', ' ').Split(' ') | Where-Object { $_ -and $_ -notin $stop } | Select-Object -Unique)
    $wb = @((("$B".ToLowerInvariant() -replace '&', ' and ') -replace '[^\p{L}\p{Nd}]+', ' ').Split(' ') | Where-Object { $_ -and $_ -notin $stop } | Select-Object -Unique)
    if ($wa.Count -eq 0 -or $wb.Count -eq 0) { return $false }
    $common = @($wa | Where-Object { $_ -in $wb }).Count
    if ($common -eq [math]::Min($wa.Count, $wb.Count)) { return $true }
    $union = $wa.Count + $wb.Count - $common
    return ($common / $union) -ge 0.6
}

# Arguments the parameter list did not take ($UnknownArgs). A single bare switch such as
# -Basketball is offered as a new genre folder (<drive>\Basketball\<Title>); the user
# confirms, or says no and the rip goes to the standard folder. Returns
# @{ Genre = "Basketball" | ""; Error = "" | message }. $KnownParams lists the real
# parameter names, so a near-miss typo (-Documentry) is pointed out before asking.
function Resolve-CustomGenreArg {
    param([object[]]$Arguments = @(), [string]$OutputDrive, [string[]]$KnownParams = @(),
          [scriptblock]$ReadInput = { param($p) Read-Host $p })
    $result = @{ Genre = ""; Error = "" }
    $Arguments = @($Arguments | Where-Object { $null -ne $_ -and "$_" -ne '' })
    if ($Arguments.Count -eq 0) { return $result }
    $switches = @($Arguments | Where-Object { "$_" -match '^-[A-Za-z][A-Za-z0-9]*:?$' })
    if ($switches.Count -ne 1 -or $Arguments.Count -ne 1) {
        $result.Error = "Unknown argument(s): $($Arguments -join ' '). Only one new folder switch (e.g. -Basketball) can be given, with no value after it."
        return $result
    }
    $name = "$($switches[0])".TrimStart('-').TrimEnd(':')
    $near = @($KnownParams | Where-Object { $_ -and (Get-EditDistance $_.ToLower() $name.ToLower()) -le 2 })
    if ($near.Count -gt 0) {
        Write-Host "`n-$name is not a rip-disc.ps1 option - did you mean -$($near[0])?" -ForegroundColor Yellow
    }
    $drive = "$OutputDrive".TrimEnd('\')
    Write-Host "`n-$name is not a built-in type. Treat it as a new folder?" -ForegroundColor Cyan
    Write-Host "  [Y] Yes - rip into $drive\$name\<Title>" -ForegroundColor White
    Write-Host "  [N] No  - rip into the standard folder ($drive\DVDs or $drive\Bluray)" -ForegroundColor White
    Write-Host "  [A] Abort" -ForegroundColor White
    while ($true) {
        $answer = & $ReadInput "Use $drive\${name}? (Y/N/A, Enter = Y)"
        if ($null -eq $answer) { return $result }
        switch ("$answer".Trim().ToUpper()) {
            { $_ -in @('Y', '') } { $result.Genre = $name; return $result }
            'N' { return $result }
            'A' { $result.Error = "Aborted."; return $result }
            default { Write-Host "Enter Y, N or A." -ForegroundColor Red }
        }
    }
}

# Levenshtein distance, for spotting a mistyped switch name.
function Get-EditDistance {
    param([string]$A, [string]$B)
    $prev = 0..$B.Length
    for ($i = 1; $i -le $A.Length; $i++) {
        $cur = @($i) + @(0) * $B.Length
        for ($j = 1; $j -le $B.Length; $j++) {
            $cost = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $cur[$j] = [math]::Min([math]::Min($prev[$j] + 1, $cur[$j - 1] + 1), $prev[$j - 1] + $cost)
        }
        $prev = $cur
    }
    return $prev[$B.Length]
}

# "Slow Dance" + "2014" -> "Slow Dance (2014)". No valid 4-digit year, no change.
# A title that already ends in a year ("Musica Surfica (2008)", as a pick-list title can)
# has that year replaced rather than doubled ("Musica Surfica (2008) (2008)", 2026-10-29).
# Parentheses are not PowerShell wildcard characters, so -Path / -like keep working.
function Get-MediaFolderName {
    param([string]$SafeTitle, [string]$Year = "")
    if ("$Year" -match '^\d{4}$') { return "$($SafeTitle -replace ' \(\d{4}\)$', '') ($Year)" }
    return $SafeTitle
}

# The year a title already carries ("Musica Surfica (2008)" -> "2008"), else "".
function Get-TitleYear {
    param([string]$Title)
    if ($Title -match ' \((\d{4})\)$') { return $Matches[1] }
    return ""
}

# The one existing folder for this title under $ParentDir, plain or with a year, so a
# Disc 2 / extras / continue-rip run lands in the folder the first run made. $null when
# there is none, or more than one (e.g. two films of the same name) - the caller decides.
# Spelling-tolerant title comparison: lower case, punctuation/spaces and a leading "the"
# ignored, then up to 2 letters different for titles of 8+ letters, 1 for 5-7, none for
# shorter ones (short titles are too easy to confuse: "Heat" / "Head"). Real case
# 2026-10-31: the Erin Brockovich DVD's label is spelled "Erin Brokovich".
function ConvertTo-LooseTitleKey {
    param([string]$Title)
    $k = "$Title".ToLowerInvariant() -replace '&', ' and '
    $k = ($k -replace '[^\p{L}\p{Nd}]+', ' ').Trim() -replace '^the\s+', ''
    return ($k -replace '\s+', '')
}

function Test-NearMissTitle {
    param([string]$A, [string]$B)
    $ka = ConvertTo-LooseTitleKey $A; $kb = ConvertTo-LooseTitleKey $B
    if (-not $ka -or -not $kb) { return $false }
    if ($ka -eq $kb) { return $true }
    # Different numbers = different titles: sequels, parts, issues ("Rocky 2" / "Rocky 3",
    # "Die Hard 2" / "Die Hard 4.0", "Issue 14" / "Issue 17"). Checked on all 1181 F: folders.
    if ((([regex]::Matches($ka, '\d+') | ForEach-Object { $_.Value }) -join ',') -ne
        (([regex]::Matches($kb, '\d+') | ForEach-Object { $_.Value }) -join ',')) { return $false }
    # One title plus extra words is a different title ("Dumb and Dumber" / "...To")
    if ($ka.StartsWith($kb) -or $kb.StartsWith($ka)) { return $false }
    $shorter = [math]::Min($ka.Length, $kb.Length)
    # 2 letters only for long titles: "Pale Rider" / "Whale Rider" are 2 apart at 9 letters
    $allowed = if ($shorter -ge 12) { 2 } elseif ($shorter -ge 5) { 1 } else { 0 }
    if ($allowed -eq 0 -or [math]::Abs($ka.Length - $kb.Length) -gt $allowed) { return $false }
    return ((Get-EditDistance -A $ka -B $kb) -le $allowed)
}

# Folders in $ParentDir spelled almost like $SafeTitle (Test-NearMissTitle on the name
# without its " (YYYY)"), excluding exact matches - Find-ExistingMediaFolder handles those.
# Lets a typo'd disc label or -title still land in the folder that already holds the film.
function Find-NearMissMediaFolders {
    param([string]$ParentDir, [string]$SafeTitle)
    if (-not $ParentDir -or -not (Test-Path -LiteralPath $ParentDir -PathType Container)) { return @() }
    $exact = '^' + [regex]::Escape($SafeTitle) + '( \(\d{4}\))?$'
    return @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch $exact -and (Test-NearMissTitle ($_.Name -replace ' \(\d{4}\)$', '') $SafeTitle) } |
        ForEach-Object { $_.Name })
}

function Find-ExistingMediaFolder {
    param([string]$ParentDir, [string]$SafeTitle)
    if (-not $ParentDir -or -not (Test-Path -LiteralPath $ParentDir -PathType Container)) { return $null }
    $pattern = '^' + [regex]::Escape($SafeTitle) + '( \(\d{4}\))?$'
    $hits = @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pattern })
    if ($hits.Count -eq 1) { return $hits[0].Name }
    return $null
}

# The folder name the old Get-SafeTitle rule gave $Title (illegal characters -> "_",
# brackets kept): "Strike Force: The Story" -> "Strike Force_ The Story".
function Get-LegacySafeTitle {
    param([string]$Title)
    return ($Title -replace '[\\/:*?"<>|]', '_').TrimEnd(' ', '.')
}

# $SafeTitle, or the old-rule name when a folder by that name (or "<name> (YYYY)") exists
# in one of $ParentDirs and none by the new name does - so titles ripped before the rule
# changed keep their folder (Disc 2 / extras / continue-rip land beside Disc 1).
function Resolve-LegacySafeTitle {
    param([string]$Title, [string]$SafeTitle, [string[]]$ParentDirs = @())
    $old = Get-LegacySafeTitle $Title
    if (-not $old -or $old -eq $SafeTitle) { return $SafeTitle }
    $dirs = @($ParentDirs | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) })
    foreach ($d in $dirs) { if (Find-ExistingMediaFolder -ParentDir $d -SafeTitle $SafeTitle) { return $SafeTitle } }
    foreach ($d in $dirs) {
        if (Find-ExistingMediaFolder -ParentDir $d -SafeTitle $old) {
            Write-Host "Using the existing folder name `"$old`" (named before titles dropped ':' etc.)" -ForegroundColor Gray
            return $old
        }
    }
    return $SafeTitle
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
