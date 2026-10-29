#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the output-drive prompt, custom genre folders (-Genre / -Basketball style
    switches), the offline check and the content-type label.

.DESCRIPTION
    Lifts the real functions out of the scripts with the AST parser, so nothing here is
    a reimplementation. The early-exit checks run the real rip-disc.ps1 with -File; they
    stop before any drive, MakeMKV or network access. No disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-OutputDriveAndGenre.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$configPath = Join-Path $repoRoot 'Load-Config.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'
$seriesPath = Join-Path $repoRoot 'SeriesEpisodes.ps1'

foreach ($p in @($configPath, $ripDiscPath, $continuePath, $seriesPath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

function Get-ScriptAst {
    param([string]$ScriptPath)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }
    return $ast
}

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $fn = (Get-ScriptAst $ScriptPath).FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

foreach ($name in @('Get-MovieParentFolder', 'Get-ContentTypeLabel', 'Test-GenreFolderName', 'Read-OutputDrive',
                    'Resolve-CustomGenreArg', 'Get-EditDistance')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
}
foreach ($name in @('Test-OnlineOnce', 'Get-YearLookup', 'Get-ContinueRipCommand')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $ripDiscPath $name)))
}
. ([scriptblock]::Create((Import-FunctionFromScript $seriesPath 'Get-SeriesDiscDbLookup')))

function New-InputQueue {
    param([object[]]$Answers)
    $queue = New-Object System.Collections.Queue
    foreach ($a in $Answers) { $queue.Enqueue($a) }
    return { param($prompt) if ($queue.Count -gt 0) { $queue.Dequeue() } else { $null } }.GetNewClosure()
}

$script:Passed = 0
$script:Failed = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) {
        $script:Passed++
        Write-Host "  PASS  $Because" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
        Write-Host "        expected: $Expected" -ForegroundColor Red
        Write-Host "        actual:   $Actual" -ForegroundColor Red
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Because)
    if ($Condition) {
        $script:Passed++
        Write-Host "  PASS  $Because" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
    }
}

$drives = @(
    [pscustomobject]@{ Letter = 'C:'; Label = 'Windows'; FreeGB = 100 }
    [pscustomobject]@{ Letter = 'E:'; Label = 'Media'; FreeGB = 500 }
    [pscustomobject]@{ Letter = 'F:'; Label = 'WD Elements'; FreeGB = 900 }
)

# ---------------------------------------------------------------------------
Write-Host "`nOutput drive prompt" -ForegroundColor Cyan

Assert-Equal 'E:' (Read-OutputDrive -Default 'E:' -Drives $drives -ReadInput (New-InputQueue @(''))) 'Enter takes the config default'
Assert-Equal 'E:' (Read-OutputDrive -Default 'e' -Drives $drives -ReadInput (New-InputQueue @())) 'end of input takes the default, normalised to "E:"'
Assert-Equal 'F:' (Read-OutputDrive -Default 'E:' -Drives $drives -ReadInput (New-InputQueue @('f'))) 'a bare letter is accepted'
Assert-Equal 'F:' (Read-OutputDrive -Default 'E:' -Drives $drives -ReadInput (New-InputQueue @('F:\'))) '"F:\" is accepted'
Assert-Equal 'F:' (Read-OutputDrive -Default 'E:' -Drives $drives -ReadInput (New-InputQueue @('Z', 'xx', 'F'))) 'a drive not in the list and junk are asked again'
Assert-Equal 'Z:' (Read-OutputDrive -Default 'E:' -Drives @() -ReadInput (New-InputQueue @('z'))) 'with no drive list, any letter is accepted'

# ---------------------------------------------------------------------------
Write-Host "`nUnknown switch as a new genre folder" -ForegroundColor Cyan

$r = Resolve-CustomGenreArg -Arguments @() -OutputDrive 'F:'
Assert-True ($r.Genre -eq '' -and $r.Error -eq '') 'no unknown arguments, nothing to do'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball') -OutputDrive 'F:' -ReadInput (New-InputQueue @('y'))
Assert-Equal 'Basketball' $r.Genre 'Y uses -Basketball as the folder name'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball') -OutputDrive 'F:' -ReadInput (New-InputQueue @(''))
Assert-Equal 'Basketball' $r.Genre 'Enter means yes'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball') -OutputDrive 'F:' -ReadInput (New-InputQueue @('n'))
Assert-True ($r.Genre -eq '' -and $r.Error -eq '') 'N falls back to the standard folder'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball') -OutputDrive 'F:' -ReadInput (New-InputQueue @())
Assert-True ($r.Genre -eq '' -and $r.Error -eq '') 'end of input falls back to the standard folder'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball') -OutputDrive 'F:' -ReadInput (New-InputQueue @('maybe', 'a'))
Assert-Equal 'Aborted.' $r.Error 'junk is asked again, A aborts'
$r = Resolve-CustomGenreArg -Arguments @('-Foo', 'bar') -OutputDrive 'F:'
Assert-True ($r.Error -match 'Unknown argument') 'a switch with a value is an error, not a folder'
$r = Resolve-CustomGenreArg -Arguments @('-Basketball', '-Golf') -OutputDrive 'F:'
Assert-True ($r.Error -match 'Only one') 'two unknown switches are an error'
$r = Resolve-CustomGenreArg -Arguments @('Jaws') -OutputDrive 'F:'
Assert-True ($r.Error -match 'Unknown argument') 'a stray value is an error'
$r = Resolve-CustomGenreArg -Arguments @('-Basket/ball') -OutputDrive 'F:'
Assert-True ($r.Error -match 'Unknown argument') 'only letters and digits make a folder switch'
Assert-True ((Get-EditDistance 'documentary' 'documentry') -le 2) 'a near-miss typo is within distance 2 of the real option'
Assert-True ((Get-EditDistance 'documentary' 'basketball') -gt 2) 'a real new name is not flagged as a typo'

Assert-True (Test-GenreFolderName 'Basketball') 'a plain name is a valid folder'
Assert-True (Test-GenreFolderName 'Rugby League') 'spaces are allowed'
Assert-True (-not (Test-GenreFolderName 'a\b')) 'a path separator is rejected'
Assert-True (-not (Test-GenreFolderName '..')) '".." is rejected'
Assert-True (-not (Test-GenreFolderName 'Bball[1]')) 'wildcard brackets are rejected'

# ---------------------------------------------------------------------------
Write-Host "`nFolders and labels" -ForegroundColor Cyan

Assert-Equal 'Basketball' (Get-MovieParentFolder -Genre 'Basketball') '-Genre is the parent folder'
Assert-Equal 'Basketball' (Get-MovieParentFolder -Genre 'Basketball' -Bluray) '-Genre wins over -Bluray, like the built-in genres'
Assert-Equal 'Surf' (Get-MovieParentFolder -Surf -Genre 'Basketball') 'a built-in genre still wins'
Assert-Equal 'DVDs' (Get-MovieParentFolder) 'plain movies still under DVDs'
Assert-Equal 'Documentary' (Get-ContentTypeLabel -GenreLabel 'Documentary') 'a -Documentary run is labelled Documentary, not Movie'
Assert-Equal 'Documentary Series' (Get-ContentTypeLabel -IsGenreSeries -GenreLabel 'Documentary' -Series) 'genre series label'
Assert-Equal 'Basketball' (Get-ContentTypeLabel -GenreLabel 'Basketball' -Bluray) 'a custom genre is its own label'
Assert-Equal 'TV Series' (Get-ContentTypeLabel -Series) 'plain series'
Assert-Equal 'Blu-ray' (Get-ContentTypeLabel -Bluray) 'Blu-ray movie'
Assert-Equal 'Movie' (Get-ContentTypeLabel) 'plain movie'

$ripText = Get-Content -LiteralPath $ripDiscPath -Raw
$contText = Get-Content -LiteralPath $continuePath -Raw
Assert-True ($ripText -notmatch "Type:\s+\`$\(if \(\`$Series\) \{ 'TV Series' \} else \{ 'Movie' \}\)") 'Discovered Metadata no longer hard-codes Movie'
Assert-True ($ripText.Contains('$finalOutputDir = "$outputDriveLetter\$Genre\$($script:MediaFolderName)"')) 'rip-disc.ps1 builds <drive>\<Genre>\<Title>'
Assert-True ($contText.Contains('$finalOutputDir = "$outputDriveLetter\$Genre\$($script:MediaFolderName)"')) 'continue-rip.ps1 builds the same folder'
Assert-True ($contText -match '\[string\]\$Genre = ""') 'continue-rip.ps1 takes -Genre'

$cmd = Get-ContinueRipCommand -Title 'Larry Bird' -RemainingStepNumbers @(2, 3, 4) -Disc 1 -Genre 'Basketball' -OutputDrive 'F:'
Assert-True ($cmd -match '-Genre "Basketball"') "the continue-rip command carries -Genre ($cmd)"
$cmd = Get-ContinueRipCommand -Title 'Jaws' -RemainingStepNumbers @(2, 3, 4) -Disc 1
Assert-True ($cmd -notmatch '-Genre') 'no -Genre when none was used'

# ---------------------------------------------------------------------------
Write-Host "`nOffline" -ForegroundColor Cyan

function Write-Log { param($m) }
function Test-InternetConnection { $script:ProbeCount++; return $false }
$script:ProbeCount = 0
$script:IsOnline = $null
Assert-True (-not (Test-OnlineOnce)) 'offline is reported'
Assert-True (-not (Test-OnlineOnce)) 'and remembered'
Assert-Equal 1 $script:ProbeCount 'the connection is probed once per run'
Assert-True ($null -eq (Get-YearLookup)) 'offline: no year search, the prompt only takes a typed year'

function Test-InternetConnection { return $true }
$script:IsOnline = $null
Assert-True (Test-OnlineOnce) 'online is reported'
Assert-True ((Get-YearLookup) -is [scriptblock]) 'online: the year search is offered'

function Get-DiscDbContentHash { param($DriveRoot) return [pscustomobject]@{ Hash = 'DA99D5A5CE83787D77675009913F3824' } }
function Find-DiscDbDisc { param($ContentHash) $script:DiscDbAsked = $true; return $null }
$script:DiscDbAsked = $false
$lookup = Get-SeriesDiscDbLookup -DriveRoot 'D:\' -Offline
Assert-True (-not $script:DiscDbAsked) 'offline: TheDiscDB is not asked'
Assert-Equal 'DA99D5A5CE83787D77675009913F3824' $lookup.ContentHash 'offline: the disc hash is still read (for continue-rip)'
Assert-True ($lookup.Message -match 'no internet') 'offline: the message says why'
$lookup = Get-SeriesDiscDbLookup -DriveRoot 'D:\'
Assert-True $script:DiscDbAsked 'online: TheDiscDB is asked as before'

# ---------------------------------------------------------------------------
Write-Host "`nrip-disc.ps1 stops early on bad arguments (real script, nothing ripped)" -ForegroundColor Cyan

function Invoke-RipDisc {
    param([string[]]$Arguments)
    $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ripDiscPath @Arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $out }
}

$run = Invoke-RipDisc @('-OutputDrive', 'F', '-Foo', 'bar')
Assert-True ($run.ExitCode -eq 1 -and $run.Output -match 'Unknown argument') 'an unknown option with a value stops with exit 1'
$run = Invoke-RipDisc @('-OutputDrive', 'F', '-Genre', 'Basketball', '-Surf')
Assert-True ($run.ExitCode -eq 1 -and $run.Output -match 'cannot be combined') '-Genre with a built-in genre stops'
$run = Invoke-RipDisc @('-OutputDrive', 'F', '-Genre', 'a\b')
Assert-True ($run.ExitCode -eq 1 -and $run.Output -match 'not a usable folder name') 'a bad -Genre folder name stops'

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
