param(
    [Parameter()]
    [string]$title = "",

    # Which step to resume from. Accepts a number (2, 3, 4) or a name
    # (handbrake, organize, open). Omit it to pick from a menu.
    [Parameter()]
    [Alias("Step", "From")]
    [string]$FromStep = "",

    [Parameter()]
    [switch]$Series,

    [Parameter()]
    [int]$Season = 0,

    [Parameter()]
    [int]$Disc = 1,

    [Parameter()]
    [string]$OutputDrive = "E:",

    [Parameter()]
    [switch]$Extras,

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

    # Accepted for command-line compatibility with rip-disc.ps1 so a failed rip
    # command can be pasted here unchanged. This script never reads the disc,
    # so both values are ignored.
    [Parameter()]
    [string]$Drive = "",

    [Parameter()]
    [int]$DriveIndex = -1,

    # Skip the "press Enter to start" confirmation.
    [Parameter()]
    [switch]$Yes,

    # Re-encode MKV files even when a matching MP4 already exists.
    [Parameter()]
    [Alias("ReEncode")]
    [switch]$Force,

    [Parameter()]
    [switch]$Help
)

# ========== LOAD CONFIG ==========
. (Join-Path $PSScriptRoot "Load-Config.ps1")

# Apply config defaults to parameters that weren't explicitly passed
if (-not $PSBoundParameters.ContainsKey('OutputDrive')) { $OutputDrive = $script:Config_DefaultOutputDrive }

# ========== GENRE SERIES (e.g. multi-disc documentary box sets) ==========
# See rip-disc.ps1 for the full explanation. Mirrored here so continuing a genre
# series rip (e.g. resuming Step 3 after an interrupted encode) organizes files
# the same way the original rip-disc.ps1 run would have.
$script:GenreLabel = if ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } else { $null }
$script:GenreFolder = if ($Documentary) { "Documentaries" } elseif ($Tutorial) { "Tutorials" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } else { $null }
$script:IsGenreSeries = $Series -and ($null -ne $script:GenreLabel)

# ========== STEP DEFINITIONS ==========
$script:AllSteps = @(
    @{ Number = 1; Key = "makemkv";   Name = "MakeMKV rip";        Description = "Rip the disc to MKV files";             Needs = "a disc in the drive"; Resumable = $false }
    @{ Number = 2; Key = "handbrake"; Name = "HandBrake encoding"; Description = "Encode the MKV files to MP4";           Needs = "MKV files in the MakeMKV temp folder"; Resumable = $true }
    @{ Number = 3; Key = "organize";  Name = "Organize files";     Description = "Rename, prefix and move the MP4 files"; Needs = "MP4 files in the output folder"; Resumable = $true }
    @{ Number = 4; Key = "open";      Name = "Open directory";     Description = "Open the finished folder in Explorer";  Needs = "the output folder to exist"; Resumable = $true }
)
$script:CompletedSteps = @()
$script:CurrentStep = $null
$script:LastWorkingDirectory = $null

function Get-Step {
    param([int]$Number)
    return $script:AllSteps | Where-Object { $_.Number -eq $Number }
}

function Get-StepByKey {
    param([string]$Key)
    return $script:AllSteps | Where-Object { $_.Key -eq $Key }
}

# Accepts "2", "handbrake", "encode", etc. Returns the canonical key or $null.
function Resolve-StepKey {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    switch -Regex ($Value.Trim()) {
        '^(1|makemkv|rip|makemkvrip)$'          { return "makemkv" }
        '^(2|handbrake|encode|encoding|hb)$'    { return "handbrake" }
        '^(3|organize|organise|rename|files)$'  { return "organize" }
        '^(4|open|openfolder|folder|explorer)$' { return "open" }
        default                                 { return $null }
    }
}

function Complete-CurrentStep {
    if ($script:CurrentStep) {
        $script:CompletedSteps += $script:CurrentStep
        Write-Host ("`n[DONE] Step {0}/4 - {1}" -f $script:CurrentStep.Number, $script:CurrentStep.Name) -ForegroundColor Green
    }
}

function Get-RemainingSteps {
    $completedNumbers = $script:CompletedSteps | ForEach-Object { $_.Number }
    return $script:AllSteps | Where-Object { $_.Number -notin $completedNumbers }
}

# "[x] 1  [>] 2  [ ] 3  [ ] 4" - completed / current / pending
function Get-StepProgressBar {
    param([int]$CurrentNumber)
    $completedNumbers = @($script:CompletedSteps | ForEach-Object { $_.Number })
    $parts = foreach ($step in $script:AllSteps) {
        if ($step.Number -eq $CurrentNumber) { "[>] $($step.Number)" }
        elseif ($completedNumbers -contains $step.Number) { "[x] $($step.Number)" }
        else { "[ ] $($step.Number)" }
    }
    return ($parts -join "  ")
}

# Printed at the top of every step so it is always obvious what is running.
function Show-StepBanner {
    param([int]$StepNumber)
    $step = Get-Step -Number $StepNumber
    $script:CurrentStep = $step
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host (" STEP {0} OF 4 - {1}" -f $step.Number, $step.Name.ToUpper()) -ForegroundColor Green
    Write-Host (" {0}" -f $step.Description) -ForegroundColor Gray
    Write-Host "========================================" -ForegroundColor Green
    Write-Host (" Progress: {0}" -f (Get-StepProgressBar -CurrentNumber $StepNumber)) -ForegroundColor DarkGray
    Write-Host ""
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
            Write-Host ("  [x] Step {0}/4  {1,-20}  {2}" -f $step.Number, $step.Name, $step.Description) -ForegroundColor Green
        }
    }

    if ($ShowRemaining) {
        $remaining = Get-RemainingSteps
        if ($remaining.Count -gt 0) {
            Write-Host "`n--- STEPS REMAINING ---" -ForegroundColor Yellow
            foreach ($step in $remaining) {
                Write-Host ("  [ ] Step {0}/4  {1,-20}  {2}" -f $step.Number, $step.Name, $step.Description) -ForegroundColor Yellow
            }
            $resumable = @($remaining | Where-Object { $_.Resumable } | Sort-Object Number)
            if ($resumable.Count -gt 0) {
                $next = $resumable[0]
                Write-Host ("`n  To pick up from here: .\continue-rip.ps1 -title `"{0}`" -FromStep {1}" -f $title, $next.Number) -ForegroundColor Cyan
            }
        }
    }
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

# ========== USAGE / HELP ==========
function Show-Usage {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " continue-rip.ps1 - resume a failed rip" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nPicks up an interrupted rip after the disc has already been read." -ForegroundColor White
    Write-Host "Run it with no arguments to be walked through the options." -ForegroundColor White

    Write-Host "`n--- STEPS ---" -ForegroundColor Cyan
    foreach ($step in $script:AllSteps) {
        if ($step.Resumable) {
            Write-Host ("  {0} / {1,-10} {2,-20}  {3}" -f $step.Number, $step.Key, $step.Name, $step.Description) -ForegroundColor White
            Write-Host ("      needs {0}" -f $step.Needs) -ForegroundColor Gray
        } else {
            Write-Host ("  {0} / {1,-10} {2,-20}  {3}" -f $step.Number, $step.Key, $step.Name, $step.Description) -ForegroundColor DarkGray
            Write-Host "      not available here - use rip-disc.ps1 to read a disc" -ForegroundColor DarkGray
        }
    }

    Write-Host "`n--- PARAMETERS ---" -ForegroundColor Cyan
    Write-Host "  -title <name>       Title as used for the original rip" -ForegroundColor White
    Write-Host "  -FromStep <2|3|4>   Step to resume from (number or name)" -ForegroundColor White
    Write-Host "  -OutputDrive <X>    Output drive letter (default $($script:Config_DefaultOutputDrive))" -ForegroundColor White
    Write-Host "  -Disc <n>           Disc number (default 1)" -ForegroundColor White
    Write-Host "  -Extras             Extras / special-features disc" -ForegroundColor White
    Write-Host "  -Series -Season <n> TV series mode" -ForegroundColor White
    Write-Host "  -Bluray             Blu-ray (subtitle handling + Bluray folder)" -ForegroundColor White
    Write-Host "  -Documentary -Tutorial -Fitness -Music -Surf   Genre folders" -ForegroundColor White
    Write-Host "  -Force              Re-encode files that already have an MP4" -ForegroundColor White
    Write-Host "                      (by default those are skipped)" -ForegroundColor Gray
    Write-Host "  -Yes                Do not ask for confirmation before starting" -ForegroundColor White
    Write-Host "  -Help               Show this text" -ForegroundColor White
    Write-Host "  -Drive / -DriveIndex   Accepted but ignored (no disc is read)" -ForegroundColor Gray

    Write-Host "`n--- EXAMPLES ---" -ForegroundColor Cyan
    Write-Host "  .\continue-rip.ps1" -ForegroundColor Yellow
    Write-Host "      Interactive - asks for everything it needs" -ForegroundColor Gray
    Write-Host "  .\continue-rip.ps1 -title `"Inception`" -FromStep 2" -ForegroundColor Yellow
    Write-Host "      Encode MKV files that were already ripped" -ForegroundColor Gray
    Write-Host "  .\continue-rip.ps1 -title `"Metal A Headbangers Journey-Disc 2`" -FromStep 3 -Documentary -OutputDrive F" -ForegroundColor Yellow
    Write-Host "      Organize already-encoded MP4 files" -ForegroundColor Gray
    Write-Host "  .\continue-rip.ps1 -title `"Fargo`" -Series -Season 1 -FromStep organize" -ForegroundColor Yellow
    Write-Host "      Series mode, names work as well as numbers" -ForegroundColor Gray
    Write-Host ""
}

if ($Help) {
    Show-Usage
    exit 0
}

# ========== INTERACTIVE PROMPTS ==========
# Only used to fill in what was not supplied on the command line. A fully
# specified command line never prompts.
# Read-Host returns $null when stdin is exhausted (redirected/non-interactive
# input), which would otherwise spin the prompt loops forever.
function Read-Answer {
    param([string]$Prompt)
    $value = Read-Host $Prompt
    if ($null -eq $value) {
        Write-Host "`nNo input available - this script cannot prompt here." -ForegroundColor Red
        Write-Host "Supply the details on the command line instead, e.g." -ForegroundColor Yellow
        Write-Host "  .\continue-rip.ps1 -title `"My Film`" -FromStep 2" -ForegroundColor Yellow
        exit 1
    }
    return $value.Trim()
}

function Read-TitleInteractive {
    Write-Host "`n--- TITLE ---" -ForegroundColor Cyan
    $tempRootPath = $script:Config_TempRoot
    $candidates = @()
    if ($tempRootPath -and (Test-Path $tempRootPath)) {
        $candidates = @(Get-ChildItem -Path $tempRootPath -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne "logs" } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 10)
    }
    if ($candidates.Count -gt 0) {
        Write-Host "Titles with files still in $tempRootPath (most recent first):" -ForegroundColor Gray
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $candidates[$i].Name) -ForegroundColor White
        }
        Write-Host "  [0] Type a title myself" -ForegroundColor Gray
    }
    while ($true) {
        $answer = Read-Answer "`nTitle (or a number from the list)"
        if ([string]::IsNullOrWhiteSpace($answer)) { continue }
        if ($candidates.Count -gt 0 -and $answer -match '^\d+$') {
            $index = [int]$answer
            if ($index -ge 1 -and $index -le $candidates.Count) { return $candidates[$index - 1].Name }
            if ($index -eq 0) {
                $typed = Read-Answer "Title"
                if ($typed) { return $typed }
                continue
            }
            Write-Host "There is no [$index] in the list." -ForegroundColor Red
            continue
        }
        return $answer
    }
}

function Read-ContentTypeInteractive {
    $types = @(
        @{ Label = "Movie (DVD)";      Switch = $null },
        @{ Label = "Movie (Blu-ray)";  Switch = "Bluray" },
        @{ Label = "TV Series";        Switch = "Series" },
        @{ Label = "Documentary";      Switch = "Documentary" },
        @{ Label = "Tutorial";         Switch = "Tutorial" },
        @{ Label = "Fitness";          Switch = "Fitness" },
        @{ Label = "Music";            Switch = "Music" },
        @{ Label = "Surf";             Switch = "Surf" }
    )
    Write-Host "`n--- CONTENT TYPE ---" -ForegroundColor Cyan
    Write-Host "This decides which folder the files end up in." -ForegroundColor Gray
    for ($i = 0; $i -lt $types.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $types[$i].Label) -ForegroundColor White
    }
    while ($true) {
        $answer = Read-Answer "`nType [1]"
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = "1" }
        if ($answer -match '^\d+$') {
            $index = [int]$answer
            if ($index -ge 1 -and $index -le $types.Count) { return $types[$index - 1].Switch }
        }
        Write-Host "Please enter a number between 1 and $($types.Count)." -ForegroundColor Red
    }
}

function Read-StepInteractive {
    Write-Host "`n--- WHICH STEP ---" -ForegroundColor Cyan
    Write-Host "Step 1 has already happened (the disc has been read)." -ForegroundColor Gray
    Write-Host ""
    foreach ($step in $script:AllSteps) {
        if ($step.Resumable) {
            Write-Host ("  [{0}] {1,-20} {2}" -f $step.Number, $step.Name, $step.Description) -ForegroundColor White
            Write-Host ("      needs {0}" -f $step.Needs) -ForegroundColor Gray
        } else {
            Write-Host ("  [{0}] {1,-20} {2}" -f $step.Number, $step.Name, $step.Description) -ForegroundColor DarkGray
            Write-Host "      not available here - use rip-disc.ps1 instead" -ForegroundColor DarkGray
        }
    }
    Write-Host "`nEverything from the chosen step onwards will run." -ForegroundColor Gray
    while ($true) {
        $answer = Read-Answer "`nContinue from step [2]"
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = "2" }
        $key = Resolve-StepKey $answer
        if ($key -eq "makemkv") {
            Write-Host "Step 1 reads the disc - this script cannot do that. Use rip-disc.ps1." -ForegroundColor Red
            continue
        }
        if ($key) { return $key }
        Write-Host "Please enter 2, 3 or 4 (or handbrake / organize / open)." -ForegroundColor Red
    }
}

function Read-YesNo {
    param([string]$Question, [bool]$Default = $false)
    $suffix = if ($Default) { "[Y/n]" } else { "[y/N]" }
    while ($true) {
        $answer = Read-Answer "$Question $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        if ($answer -match '^(y|yes)$') { return $true }
        if ($answer -match '^(n|no)$') { return $false }
        Write-Host "Please answer y or n." -ForegroundColor Red
    }
}

# ========== RESOLVE PARAMETERS ==========
$stepKey = Resolve-StepKey $FromStep

if ($PSBoundParameters.ContainsKey('Drive') -or $PSBoundParameters.ContainsKey('DriveIndex')) {
    Write-Host "`nNote: -Drive/-DriveIndex are ignored here - continue-rip.ps1 never reads the disc." -ForegroundColor DarkGray
}

$needsPrompting = [string]::IsNullOrWhiteSpace($title) -or ($null -eq $stepKey) -or ($stepKey -eq "makemkv")

if ($needsPrompting) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " CONTINUE RIP - resume an interrupted rip" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Run with -Help to see all parameters and examples." -ForegroundColor Gray

    if (-not [string]::IsNullOrWhiteSpace($FromStep) -and $null -eq $stepKey) {
        Write-Host "`n'$FromStep' is not a step this script recognises." -ForegroundColor Yellow
    } elseif ($stepKey -eq "makemkv") {
        Write-Host "`nStep 1 (MakeMKV rip) reads the disc - use rip-disc.ps1 for that." -ForegroundColor Yellow
        $stepKey = $null
    }

    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = Read-TitleInteractive
    } else {
        Write-Host "`nTitle: $title" -ForegroundColor White
    }

    # Content type: only ask if no type switch was supplied.
    $anyTypeSwitch = $Series -or $Bluray -or $Documentary -or $Tutorial -or $Fitness -or $Music -or $Surf
    if (-not $anyTypeSwitch) {
        $chosenType = Read-ContentTypeInteractive
        switch ($chosenType) {
            "Bluray"      { $Bluray = $true }
            "Series"      { $Series = $true }
            "Documentary" { $Documentary = $true }
            "Tutorial"    { $Tutorial = $true }
            "Fitness"     { $Fitness = $true }
            "Music"       { $Music = $true }
            "Surf"        { $Surf = $true }
        }
    }

    if ($Series) {
        if (-not $PSBoundParameters.ContainsKey('Season')) {
            $answer = Read-Answer "`nSeason number (Enter for none) [$Season]"
            if ($answer -match '^\d+$') { $Season = [int]$answer }
        }
        if (-not $PSBoundParameters.ContainsKey('Disc')) {
            $answer = Read-Answer "Disc number [$Disc]"
            if ($answer -match '^\d+$') { $Disc = [int]$answer }
        }
    } else {
        if (-not $PSBoundParameters.ContainsKey('Extras') -and -not $Extras) {
            $Extras = Read-YesNo -Question "`nIs this an extras / special-features disc?" -Default $false
        }
        if (-not $Extras -and -not $PSBoundParameters.ContainsKey('Disc')) {
            $answer = Read-Answer "Disc number [$Disc]"
            if ($answer -match '^\d+$') { $Disc = [int]$answer }
        }
    }

    if ($null -eq $stepKey) {
        $stepKey = Read-StepInteractive
    }
}

$FromStep = $stepKey
$startStep = Get-StepByKey -Key $stepKey
$StartFromStepNumber = $startStep.Number

# Mark steps before the starting point as "skipped/assumed complete"
for ($i = 1; $i -lt $StartFromStepNumber; $i++) {
    $script:CompletedSteps += Get-Step -Number $i
}

# ========== CLOSE BUTTON PROTECTION ==========
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
    [Win32.ConsoleCloseProtection]::EnableMenuItem($script:ConsoleSystemMenu, 0xF060, 0x00000001) | Out-Null
}

function Enable-ConsoleClose {
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

    $driveLetter = [System.IO.Path]::GetPathRoot($Path)
    if (-not $driveLetter) {
        return @{ Ready = $false; Drive = "Unknown"; Message = "Could not determine drive letter from path: $Path" }
    }

    $driveDisplay = $driveLetter.TrimEnd('\')

    try {
        $drive = Get-PSDrive -Name $driveDisplay.TrimEnd(':') -ErrorAction Stop
        if ($drive) {
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

# ========== CONFIGURATION ==========
$tempRoot = $script:Config_TempRoot
# MakeMKV temp directory - use subdirectory for multi-disc and extras rips
if ($Extras) {
    $makemkvOutputDir = "$tempRoot\$title\Extras"
} elseif ($Series -and $Season -gt 0) {
    $makemkvOutputDir = "$tempRoot\$title\Season$Season\Disc$Disc"
} else {
    $makemkvOutputDir = "$tempRoot\$title\Disc$Disc"
}

# Normalize output drive letter
$outputDriveLetter = if ($OutputDrive -match ':$') { $OutputDrive } else { "${OutputDrive}:" }

# Build final output directory path
if ($script:IsGenreSeries) {
    $genreSeriesBaseDir = "$outputDriveLetter\$($script:GenreFolder)\$title"
    if ($Season -gt 0) {
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $genreSeriesSeasonDir = Join-Path $genreSeriesBaseDir $seasonFolder
    } else {
        $seasonTag = $null
        $genreSeriesSeasonDir = $genreSeriesBaseDir
    }
    $finalOutputDir = Join-Path $genreSeriesSeasonDir "Disc$Disc"
} elseif ($Documentary) {
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
        $seasonTag = "S{0:D2}" -f $Season
        $seasonFolder = "Season $Season"
        $seriesSeasonDir = Join-Path $seriesBaseDir $seasonFolder
    } else {
        $seasonTag = $null
        $seriesSeasonDir = $seriesBaseDir
    }
    # Use per-disc subdirectory to isolate concurrent disc rips (prevents rename conflicts)
    $finalOutputDir = Join-Path $seriesSeasonDir "Disc$Disc"
} elseif ($Bluray) {
    $finalOutputDir = "$outputDriveLetter\Bluray\$title"
} else {
    $finalOutputDir = "$outputDriveLetter\DVDs\$title"
}

# Extras: encode directly into extras subdirectory of the title folder
if ($Extras -and -not $Series) {
    $finalOutputDir = Join-Path $finalOutputDir "extras"
}

$handbrakePath = $script:Config_HandBrakePath

# ========== LOGGING SETUP ==========
$logDir = Join-Path $tempRoot "logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDiscLabel = if ($Extras) { "extras" } else { "disc${Disc}" }
$script:LogFile = Join-Path $logDir "${title}_${logDiscLabel}_continue_${logTimestamp}.log"

Write-Log "========== CONTINUE SESSION STARTED =========="
Write-Log "Title: $title"
Write-Log "Continue from: Step $StartFromStepNumber ($FromStep)"
Write-Log "Type: $(if ($script:IsGenreSeries) { "$($script:GenreLabel) Series" } elseif ($Documentary) { 'Documentary' } elseif ($Tutorial) { 'Tutorial' } elseif ($Fitness) { 'Fitness' } elseif ($Music) { 'Music' } elseif ($Surf) { 'Surf' } elseif ($Series) { 'TV Series' } elseif ($Bluray) { 'Blu-ray' } else { 'Movie' })"
Write-Log "Disc: $Disc$(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1 -and -not $Series) { ' (Special Features)' })"
if ($Series -and $Season -gt 0) {
    Write-Log "Season: $Season"
}
Write-Log "Output Drive: $outputDriveLetter"
Write-Log "MakeMKV Output: $makemkvOutputDir"
Write-Log "Final Output: $finalOutputDir"
Write-Log "Log file: $($script:LogFile)"

function Stop-WithError {
    param([string]$Step, [string]$Message)

    $host.UI.RawUI.WindowTitle = "$($host.UI.RawUI.WindowTitle) - ERROR"

    Write-Log "========== ERROR =========="
    Write-Log "Failed at: $Step"
    Write-Log "Message: $Message"
    if ($script:CompletedSteps.Count -gt 0) {
        Write-Log "Completed steps: $(($script:CompletedSteps | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    }
    $remaining = Get-RemainingSteps
    if ($remaining.Count -gt 0) {
        Write-Log "Remaining steps: $(($remaining | ForEach-Object { "Step $($_.Number): $($_.Name)" }) -join ', ')"
    }

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "`nProcessing: $(Get-TitleSummary)" -ForegroundColor White
    Write-Host "`nError at: $Step" -ForegroundColor Red
    Write-Host "Message: $Message" -ForegroundColor Red

    Show-StepsSummary -ShowRemaining

    $directoryToOpen = $null
    if ($script:LastWorkingDirectory -and (Test-Path $script:LastWorkingDirectory)) {
        $directoryToOpen = $script:LastWorkingDirectory
    } elseif (Test-Path $makemkvOutputDir) {
        $directoryToOpen = $makemkvOutputDir
    } elseif (Test-Path $finalOutputDir) {
        $directoryToOpen = $finalOutputDir
    }

    if ($directoryToOpen) {
        Write-Host "`n--- OPENING DIRECTORY ---" -ForegroundColor Cyan
        Write-Host "Opening: $directoryToOpen" -ForegroundColor Yellow
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

    Write-Host "`nLog file: $($script:LogFile)" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Red
    Enable-ConsoleClose
    exit 1
}

$contentType = if ($script:IsGenreSeries) { "$($script:GenreLabel) Series" } elseif ($Documentary) { "Documentary" } elseif ($Tutorial) { "Tutorial" } elseif ($Fitness) { "Fitness" } elseif ($Music) { "Music" } elseif ($Surf) { "Surf" } elseif ($Series) { "TV Series" } elseif ($Bluray) { "Blu-ray" } else { "Movie" }
$isMainFeatureDisc = (-not $Series) -and ($Disc -eq 1) -and (-not $Extras)
$extrasDir = Join-Path $finalOutputDir "extras"

# ========== WINDOW TITLE ==========
if ($Series) {
    $windowTitle = "$title"
    if ($Season -gt 0) { $windowTitle += " S$Season" }
    $windowTitle += " Disc $Disc"
} else {
    $windowTitle = "$title"
    if ($Extras -or $Disc -gt 1) { $windowTitle += "-extras" }
}
$windowTitle += " - CONTINUE"
$host.UI.RawUI.WindowTitle = $windowTitle

# ========== VALIDATION ==========
# Prints the exact command that reproduces this run, so it can be re-used or tweaked.
function Get-EquivalentCommand {
    $parts = @(".\continue-rip.ps1", "-title `"$title`"", "-FromStep $StartFromStepNumber")
    if ($Series)      { $parts += "-Series" }
    if ($Season -gt 0){ $parts += "-Season $Season" }
    if ($Disc -ne 1)  { $parts += "-Disc $Disc" }
    if ($Extras)      { $parts += "-Extras" }
    if ($Bluray)      { $parts += "-Bluray" }
    if ($Documentary) { $parts += "-Documentary" }
    if ($Tutorial)    { $parts += "-Tutorial" }
    if ($Fitness)     { $parts += "-Fitness" }
    if ($Music)       { $parts += "-Music" }
    if ($Surf)        { $parts += "-Surf" }
    if ($Force)       { $parts += "-Force" }
    $parts += "-OutputDrive $($outputDriveLetter.TrimEnd(':'))"
    return ($parts -join " ")
}

# Explain what was expected and, where possible, which step the files on disk
# actually match. Returns $false so the caller can offer another step.
function Write-PrerequisiteFailure {
    param([int]$StepNumber, [string]$Message, [string[]]$Hints = @())
    $step = Get-Step -Number $StepNumber
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host " CANNOT START AT STEP $StepNumber - $($step.Name)" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host " $Message" -ForegroundColor Red
    Write-Host " This step needs $($step.Needs)." -ForegroundColor Yellow
    foreach ($hint in $Hints) {
        Write-Host " $hint" -ForegroundColor Cyan
    }
    Write-Host "========================================" -ForegroundColor Red
    return $false
}

# Every step, what it does, what it needs, and which one is currently selected.
function Show-StepOptions {
    param([int]$SelectedNumber)
    Write-Host "`n--- STEPS AVAILABLE ---" -ForegroundColor Cyan
    foreach ($step in $script:AllSteps) {
        $marker = if ($step.Number -eq $SelectedNumber) { "-->" } else { "   " }
        if (-not $step.Resumable) {
            Write-Host ("{0} [{1}] {2,-20} {3}" -f $marker, $step.Number, $step.Name, $step.Description) -ForegroundColor DarkGray
            Write-Host "        not available here - use rip-disc.ps1 to read a disc" -ForegroundColor DarkGray
            continue
        }
        if ($step.Number -eq $SelectedNumber) {
            Write-Host ("{0} [{1}] {2,-20} {3}" -f $marker, $step.Number, $step.Name, $step.Description) -ForegroundColor Yellow
            Write-Host ("        STARTS HERE - needs {0}" -f $step.Needs) -ForegroundColor Yellow
        } elseif ($step.Number -lt $SelectedNumber) {
            Write-Host ("{0} [{1}] {2,-20} {3}" -f $marker, $step.Number, $step.Name, $step.Description) -ForegroundColor DarkGray
            Write-Host "        SKIPPED - assumed already done" -ForegroundColor DarkGray
        } else {
            Write-Host ("{0} [{1}] {2,-20} {3}" -f $marker, $step.Number, $step.Name, $step.Description) -ForegroundColor White
            Write-Host "        runs after the step above" -ForegroundColor Gray
        }
    }
}

# Checks the starting step can actually run. Returns $true/$false.
function Test-StepPrerequisites {
    param([int]$StepNumber)

    if ($StepNumber -eq 2) {
        # Need MKV files in makemkvOutputDir
        $hints = @()
        if ((Test-Path $finalOutputDir) -and @(Get-ChildItem -Path $finalOutputDir -Filter "*.mp4" -ErrorAction SilentlyContinue).Count -gt 0) {
            $hints += "There ARE encoded MP4 files in $finalOutputDir - did you mean step 3 (organize)?"
        }
        if (!(Test-Path $makemkvOutputDir)) {
            return (Write-PrerequisiteFailure -StepNumber $StepNumber -Message "The MakeMKV folder does not exist: $makemkvOutputDir" -Hints $hints)
        }
        $mkvFiles = Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue
        if ($null -eq $mkvFiles -or $mkvFiles.Count -eq 0) {
            return (Write-PrerequisiteFailure -StepNumber $StepNumber -Message "No MKV files in: $makemkvOutputDir" -Hints $hints)
        }
        Write-Host "OK - found $($mkvFiles.Count) MKV file(s):" -ForegroundColor Green
        foreach ($mkv in $mkvFiles) {
            Write-Host "  - $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB)" -ForegroundColor Gray
        }
        # Say up front which files will be encoded and which are already done.
        $alreadyEncoded = @($mkvFiles | Where-Object { Test-Path (Join-Path $finalOutputDir ($_.BaseName + ".mp4")) })
        if ($alreadyEncoded.Count -gt 0) {
            if ($Force) {
                Write-Host "`n-Force given: all $($mkvFiles.Count) file(s) will be encoded, including" -ForegroundColor Yellow
                Write-Host "$($alreadyEncoded.Count) that already have an MP4 and will be overwritten:" -ForegroundColor Yellow
            } else {
                Write-Host "`n$($alreadyEncoded.Count) of these already have an MP4 and will be SKIPPED" -ForegroundColor Yellow
                Write-Host "(run with -Force to re-encode them anyway):" -ForegroundColor Yellow
            }
            foreach ($done in $alreadyEncoded) {
                Write-Host "  - $($done.BaseName).mp4" -ForegroundColor Yellow
            }
            $toEncodeCount = if ($Force) { $mkvFiles.Count } else { $mkvFiles.Count - $alreadyEncoded.Count }
            Write-Host "Files to encode this run: $toEncodeCount" -ForegroundColor Green
        }
        return $true
    }

    if ($StepNumber -eq 3) {
        # Need MP4 files in finalOutputDir
        $hints = @()
        if ((Test-Path $makemkvOutputDir) -and @(Get-ChildItem -Path $makemkvOutputDir -Filter "*.mkv" -ErrorAction SilentlyContinue).Count -gt 0) {
            $hints += "There ARE un-encoded MKV files in $makemkvOutputDir - did you mean step 2 (handbrake)?"
        }
        if (!(Test-Path $finalOutputDir)) {
            return (Write-PrerequisiteFailure -StepNumber $StepNumber -Message "The output folder does not exist: $finalOutputDir" -Hints $hints)
        }
        $mp4Files = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4" -ErrorAction SilentlyContinue
        if ($null -eq $mp4Files -or $mp4Files.Count -eq 0) {
            return (Write-PrerequisiteFailure -StepNumber $StepNumber -Message "No MP4 files in: $finalOutputDir" -Hints $hints)
        }
        Write-Host "OK - found $($mp4Files.Count) MP4 file(s) to organize:" -ForegroundColor Green
        foreach ($mp4 in $mp4Files) {
            Write-Host "  - $($mp4.Name) ($([math]::Round($mp4.Length/1GB, 2)) GB)" -ForegroundColor Gray
        }
        return $true
    }

    if ($StepNumber -eq 4) {
        # Just need finalOutputDir to exist
        if (!(Test-Path $finalOutputDir)) {
            return (Write-PrerequisiteFailure -StepNumber $StepNumber -Message "The output folder does not exist: $finalOutputDir")
        }
        Write-Host "OK - output folder exists: $finalOutputDir" -ForegroundColor Green
        return $true
    }

    return $false
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " CONTINUE RIP" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host (" Title         : {0}" -f $title) -ForegroundColor White
Write-Host (" Type          : {0}" -f $contentType) -ForegroundColor White
if ($Series) {
    if ($Season -gt 0) {
        Write-Host (" Season        : {0} ({1})" -f $Season, $seasonTag) -ForegroundColor White
    }
    Write-Host (" Disc          : {0}" -f $Disc) -ForegroundColor White
} else {
    Write-Host (" Disc          : {0}{1}" -f $Disc, $(if ($Extras) { ' (Extras)' } elseif ($Disc -gt 1) { ' (Special Features)' } else { '' })) -ForegroundColor White
}
Write-Host (" Source folder : {0}" -f $makemkvOutputDir) -ForegroundColor White
Write-Host (" Output folder : {0}" -f $finalOutputDir) -ForegroundColor White
Write-Host (" Log file      : {0}" -f $script:LogFile) -ForegroundColor White

# Show the full list of steps, check the chosen one can actually run, and let the
# step be changed at the prompt rather than re-running the whole command.
while ($true) {
    $startStep = Get-StepByKey -Key $FromStep
    $StartFromStepNumber = $startStep.Number

    # Steps before the starting point count as "skipped/assumed complete"
    $script:CompletedSteps = @()
    for ($i = 1; $i -lt $StartFromStepNumber; $i++) {
        $script:CompletedSteps += Get-Step -Number $i
    }

    Show-StepOptions -SelectedNumber $StartFromStepNumber

    Write-Host "`n--- CHECKING PREREQUISITES ---" -ForegroundColor Cyan
    $ready = Test-StepPrerequisites -StepNumber $StartFromStepNumber

    if ($ready) {
        Write-Host "`nCommand for this run:" -ForegroundColor DarkGray
        Write-Host "  $(Get-EquivalentCommand)" -ForegroundColor DarkGray
    }

    if ($Yes) {
        if (-not $ready) { exit 1 }
        Write-Host "`nStarting (-Yes was given)..." -ForegroundColor Green
        break
    }

    # Read-Answer exits rather than starting an encode on an unanswered prompt.
    if ($ready) {
        $answer = Read-Answer "`nPress Enter to start at step $StartFromStepNumber, type another step number (2/3/4), or Ctrl+C to abort"
        if ([string]::IsNullOrWhiteSpace($answer)) { break }
    } else {
        $answer = Read-Answer "`nType another step number (2/3/4) to try a different step, or Ctrl+C to abort"
        if ([string]::IsNullOrWhiteSpace($answer)) {
            Write-Host "Step $StartFromStepNumber cannot run here - choose another step, or press Ctrl+C to abort." -ForegroundColor Red
            continue
        }
    }

    $newKey = Resolve-StepKey $answer
    if ($newKey -eq "makemkv") {
        Write-Host "Step 1 reads the disc - use rip-disc.ps1 for that." -ForegroundColor Red
        continue
    }
    if (-not $newKey) {
        Write-Host "Please enter 2, 3 or 4 (or handbrake / organize / open)." -ForegroundColor Red
        continue
    }
    if ($newKey -eq $FromStep) {
        if ($ready) { break }
        continue
    }
    $FromStep = $newKey
    Write-Log "Starting step changed at the prompt to: $FromStep"
}
Disable-ConsoleClose

# ========== STEP 2: ENCODE WITH HANDBRAKE ==========
if ($StartFromStepNumber -le 2) {
    Show-StepBanner -StepNumber 2
    $script:LastWorkingDirectory = $finalOutputDir
    Write-Log "STEP 2/4: Starting HandBrake encoding..."

    # Check if destination drive is ready
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

    # ========== BLU-RAY MODE GUARD ==========
    # A single title cannot be bigger than the disc it came from, and DVD-9 (dual
    # layer) tops out at ~8.5 GB. Any MKV above that is therefore proof of a Blu-ray
    # source, regardless of what was passed on the command line.
    #
    # Without -Bluray the DVD subtitle branch runs (--all-subtitles
    # --subtitle-burned=none), and Blu-ray PGS tracks get burned into the picture
    # anyway despite that flag - see PR #67. Burned-in subtitles cannot be removed
    # afterwards, so this check runs before any encoding starts.
    #
    # Checked per-file rather than on the total: MakeMKV often emits the same feature
    # as more than one title, so a DVD's files can legitimately sum past 8.5 GB.
    if (-not $Bluray) {
        $dvd9CapacityBytes = 8.5GB
        $oversizedMkvs = @($mkvFiles | Where-Object { $_.Length -gt $dvd9CapacityBytes } | Sort-Object Length -Descending)

        if ($oversizedMkvs.Count -gt 0) {
            $biggestMkv = $oversizedMkvs[0]
            $biggestGB = [math]::Round($biggestMkv.Length / 1GB, 2)

            Write-Host "`n========================================" -ForegroundColor Yellow
            Write-Host "BLU-RAY DETECTED, BUT -Bluray WAS NOT PASSED" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "Largest MKV: $($biggestMkv.Name) ($biggestGB GB)" -ForegroundColor White
            Write-Host "A DVD cannot hold a single title that large (DVD-9 limit is 8.5 GB)." -ForegroundColor White
            Write-Host "`nIn DVD mode HandBrake burns the PGS subtitle tracks into the picture." -ForegroundColor Red
            Write-Host "That cannot be undone once encoding has finished." -ForegroundColor Red
            Write-Log "Blu-ray guard: $($biggestMkv.Name) is $biggestGB GB (over 8.5 GB) but -Bluray was not passed"

            if ($Yes) {
                # Non-interactive: take the corrective, non-destructive option.
                $Bluray = $true
                Write-Host "`n-Yes given - enabling Blu-ray subtitle handling automatically." -ForegroundColor Green
                Write-Log "Blu-ray guard: -Yes given, -Bluray enabled automatically"
            } else {
                Write-Host "`n  [Enter] Switch to Blu-ray subtitle handling (recommended)" -ForegroundColor Green
                Write-Host "  [d]     Continue in DVD mode anyway - subtitles will be burned in" -ForegroundColor Gray
                Write-Host "  [a]     Abort and leave the MKV files alone" -ForegroundColor Gray
                $blurayChoice = (Read-Answer "`nChoice").ToLowerInvariant()

                if ($blurayChoice -eq 'a') {
                    Write-Host "`nAborted before encoding. The MKV files are untouched at:" -ForegroundColor Yellow
                    Write-Host "  $makemkvOutputDir" -ForegroundColor White
                    Write-Host "`nRe-run with -Bluray when you are ready:" -ForegroundColor Yellow
                    Write-Host "  .\continue-rip.ps1 -title `"$title`" -FromStep 2 -Bluray" -ForegroundColor Cyan
                    Write-Log "Blu-ray guard: aborted before encoding - MKVs left at $makemkvOutputDir"
                    exit 1
                } elseif ($blurayChoice -eq 'd') {
                    Write-Host "`nContinuing in DVD mode - subtitles will be burned in." -ForegroundColor Yellow
                    Write-Log "Blu-ray guard: continuing in DVD mode despite Blu-ray-sized titles"
                } else {
                    $Bluray = $true
                    Write-Host "`nBlu-ray subtitle handling enabled for the rest of this run." -ForegroundColor Green
                    Write-Log "Blu-ray guard: -Bluray enabled automatically for the remainder of this run"
                    # $finalOutputDir was resolved before this step, so it is deliberately
                    # not re-routed here - it already exists and may already hold MP4s.
                    Write-Host "Output directory is unchanged: $finalOutputDir" -ForegroundColor Gray
                    Write-Log "Blu-ray guard: output directory left as $finalOutputDir"
                }
            }
        }
    }

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

    # ========== SKIP ALREADY-ENCODED FILES ==========
    # Resuming usually means encoding stopped part way through, so files that
    # already have an MP4 are left alone. -Force re-encodes everything.
    $filesToEncode = @($mkvFiles)
    if ($Force) {
        Write-Host "`n-Force given - re-encoding every file, including any that already have an MP4." -ForegroundColor Yellow
        Write-Log "-Force given - re-encoding all $($filesToEncode.Count) file(s)"
    } else {
        $alreadyDone = @($mkvFiles | Where-Object { Test-Path (Join-Path $finalOutputDir ($_.BaseName + ".mp4")) })
        if ($alreadyDone.Count -gt 0) {
            $filesToEncode = @($mkvFiles | Where-Object { -not (Test-Path (Join-Path $finalOutputDir ($_.BaseName + ".mp4"))) })
            Write-Host "`nSkipping $($alreadyDone.Count) file(s) that already have an MP4:" -ForegroundColor Yellow
            foreach ($done in $alreadyDone) {
                $existingMp4 = Get-Item -LiteralPath (Join-Path $finalOutputDir ($done.BaseName + ".mp4"))
                Write-Host ("  - {0}.mp4 ({1} GB)" -f $done.BaseName, [math]::Round($existingMp4.Length/1GB, 2)) -ForegroundColor Gray
                Write-Log "Skipping (already encoded): $($done.Name) -> $($done.BaseName).mp4 ($([math]::Round($existingMp4.Length/1GB, 2)) GB)"
            }
            Write-Host "This assumes those MP4s are complete - run with -Force to re-encode them anyway." -ForegroundColor Gray
        }
    }

    if ($filesToEncode.Count -eq 0) {
        Write-Host "`nNothing to encode - every MKV already has an MP4." -ForegroundColor Green
        Write-Log "STEP 2/4: Nothing to encode - all files already have an MP4"
    } else {
        Write-Host "`n$($filesToEncode.Count) file(s) to encode" -ForegroundColor Green
    }

    # ========== GENERATE RECOVERY SCRIPT ==========
    $recoveryScriptPath = $null
    if ($filesToEncode.Count -gt 0) {
        $safeTitle = $title -replace '[\\/:*?"<>|]', '_'
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
        foreach ($mkv in $filesToEncode) {
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
    }

    $fileCount = 0
    foreach ($mkv in $filesToEncode) {
        $fileCount++
        $inputFile = $mkv.FullName
        $outputFile = Join-Path $finalOutputDir ($mkv.BaseName + ".mp4")

        Write-Host "`n--- Encoding file $fileCount of $($filesToEncode.Count) ---" -ForegroundColor Cyan
        Write-Host "Input:  $($mkv.Name)" -ForegroundColor White
        Write-Host "Output: $($mkv.BaseName).mp4" -ForegroundColor White
        Write-Host "Size:   $([math]::Round($mkv.Length/1GB, 2)) GB" -ForegroundColor White
        Write-Log "Encoding file $fileCount of $($filesToEncode.Count): $($mkv.Name) ($([math]::Round($mkv.Length/1GB, 2)) GB)"

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
            Write-Host "`nEncoding complete: $($mkv.Name)" -ForegroundColor Green
            Write-Host "Output size: $([math]::Round($encodedSize/1GB, 2)) GB" -ForegroundColor White
            Write-Log "Encoded: $($mkv.Name) -> $($mkv.BaseName).mp4 ($([math]::Round($encodedSize/1GB, 2)) GB)"
        } else {
            Stop-WithError -Step "STEP 2/4: HandBrake encoding" -Message "Output file not created for $($mkv.Name)"
        }
    }
    Complete-CurrentStep
    Write-Log "STEP 2/4: HandBrake encoding complete - $fileCount file(s) encoded"

    # Delete recovery script after successful encoding (none is written when
    # there was nothing to encode)
    if ($recoveryScriptPath -and (Test-Path $recoveryScriptPath)) {
        Remove-Item $recoveryScriptPath -Force
        Write-Host "Recovery script deleted (encoding successful)" -ForegroundColor Gray
        Write-Log "Recovery script deleted: $recoveryScriptPath"
    }

    # Wait for file handles
    Write-Host "`nWaiting for file handles to be released..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    Write-Host "File handle wait complete" -ForegroundColor Green

    # Delete MakeMKV temp directory after successful encode
    Write-Host "`nChecking for successful encodes..." -ForegroundColor Yellow
    $encodedFiles = Get-ChildItem -Path $finalOutputDir -Filter "*.mp4"
    $script:EncodedFilesTooSmall = $false
    if ($encodedFiles.Count -gt 0) {
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
            Start-Process explorer.exe -ArgumentList "shell:RecycleBinFolder"
            Write-Host "Opened Recycle Bin for review" -ForegroundColor Yellow
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
}

# ========== STEP 3: RENAME AND ORGANIZE ==========
if ($StartFromStepNumber -le 3) {
    Show-StepBanner -StepNumber 3
    $script:LastWorkingDirectory = $finalOutputDir
    Write-Log "STEP 3/4: Organizing files..."

    # Everything below renames/moves files in the CURRENT directory. If this cd
    # silently fails, that current directory is wherever the script was launched
    # from - which is how a run once renamed the files in the repo it was run
    # from. Fail loudly instead, and confirm we actually landed in the target.
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

    # Delete image files
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
        # the per-disc subdirectory into the shared title (or Season) folder. See
        # rip-disc.ps1 for the full explanation - this mirrors it for resumed rips.
        Write-Host "`nNumbering $($script:GenreLabel.ToLower()) series episodes..." -ForegroundColor Yellow
        $seasonTag = if ($Season -gt 0) { "S{0:D2}" -f $Season } else { "" }

        $genreSeriesTargetDir = Split-Path $finalOutputDir -Parent

        if ($PSBoundParameters.ContainsKey('StartEpisode')) {
            $nextEpisode = $StartEpisode
            Write-Host "Starting at episode $nextEpisode (explicit -StartEpisode)" -ForegroundColor Gray
            Write-Log "Genre series: starting at episode $nextEpisode (explicit -StartEpisode)"
        } else {
            $episodeNumberPattern = if ($seasonTag) { "$seasonTag`E(\d+)" } else { "-E(\d+)\." }
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

        if (!(Test-Path $genreSeriesTargetDir)) {
            New-Item -ItemType Directory -Path $genreSeriesTargetDir -Force | Out-Null
        }

        foreach ($file in $episodeFiles) {
            $episodeTag = if ($seasonTag) { "$seasonTag`E{0:D2}" -f $nextEpisode } else { "E{0:D2}" -f $nextEpisode }
            $candidateName = "$title-$episodeTag$($file.Extension)"
            $uniquePath = Get-UniqueFilePath -DestDir $genreSeriesTargetDir -FileName $candidateName
            $finalName = [System.IO.Path]::GetFileName($uniquePath)
            Write-Host "  $($file.Name) -> $finalName" -ForegroundColor Gray

            $maxRetries = 5
            $retryDelay = 3
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
        $seasonTag = if ($Season -gt 0) { "S{0:D2}" -f $Season } else { "" }
        $discTag = "D$Disc"
        $prefix = "$title-$seasonTag-$discTag"

        $episodeFiles = Get-ChildItem -File | Where-Object {
            $_.Extension -match '\.(mp4|mkv)$'
        } | Sort-Object Name

        foreach ($file in $episodeFiles) {
            $newName = "$prefix-$($file.Name)"
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
        }
        Write-Host "Renamed $($episodeFiles.Count) file(s)" -ForegroundColor Green
        Write-Log "Prefixed $($episodeFiles.Count) file(s) with $prefix"

        # Keep files in disc subdirectory (Jellyfin scans recursively)
        Write-Host "Episodes kept in disc directory: $finalOutputDir" -ForegroundColor Green
        Write-Log "Episodes kept in disc directory: $finalOutputDir"
    } else {
        # ========== MOVIE MODE ==========
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
                        $newName = $dirName + "-" + $file.Name.Substring($dirName.Length + 1)
                    } else {
                        $newName = $dirName + "-" + $file.Name
                    }
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
                    if ($file.Name -like ("$title" + "_*")) {
                        $newName = "$title-" + $file.Name.Substring($title.Length + 1)
                    } else {
                        $newName = "$title-" + $file.Name
                    }
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

        # Handle extras folder
        if ($isMainFeatureDisc) {
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
            Write-Host "`nMoving special features to extras folder..." -ForegroundColor Yellow

            if (!(Test-Path $extrasDir)) {
                Write-Host "Creating extras directory..." -ForegroundColor Yellow
                New-Item -ItemType Directory -Path $extrasDir | Out-Null
                Write-Host "Extras directory created" -ForegroundColor Green
            } else {
                Write-Host "Extras directory already exists" -ForegroundColor Gray
            }

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
}

# ========== STEP 4: OPEN DIRECTORY ==========
if ($StartFromStepNumber -le 4) {
    Show-StepBanner -StepNumber 4
    Write-Log "STEP 4/4: Opening directory..."
    Write-Host "Opening: $finalOutputDir" -ForegroundColor Yellow
    start $finalOutputDir
    Complete-CurrentStep
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nProcessed: $(Get-TitleSummary)" -ForegroundColor White
Write-Host "Final location: $finalOutputDir" -ForegroundColor White

Show-StepsSummary

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
Show-CoffeeBadge

Write-Log "========== CONTINUE SESSION COMPLETE =========="
Write-Log "Final location: $finalOutputDir"
if (-not $script:EncodedFilesTooSmall) {
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
