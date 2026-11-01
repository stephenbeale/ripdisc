#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for media server library folders: the drive scan and suggestions (setup.ps1
    step 5) and the rip-time library prompt that follows the output-drive prompt.

.DESCRIPTION
    Lifts the real functions out of the scripts with the AST parser. The drive scan runs
    against a temp folder laid out like a media drive. No disc, MakeMKV or HandBrake.

.EXAMPLE
    .\tests\Test-MediaLibraries.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$configPath = Join-Path $repoRoot 'Load-Config.ps1'
$librariesPath = Join-Path $repoRoot 'MediaLibraries.ps1'
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$setupPath = Join-Path $repoRoot 'setup.ps1'

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ScriptPath has parse errors" }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FunctionName }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

foreach ($name in @('Test-SystemFolderName', 'Test-LikelyLibraryFolder', 'Get-DriveLibraryCandidates')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $librariesPath $name)))
}
foreach ($name in @('ConvertTo-LibraryChoice', 'Read-LibraryChoice', 'Test-GenreFolderName')) {
    . ([scriptblock]::Create((Import-FunctionFromScript $configPath $name)))
}

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
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because`n        expected: $Expected`n        actual:   $Actual" -ForegroundColor Red }
}
function Assert-True {
    param([bool]$Condition, [string]$Because)
    if ($Condition) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  FAIL  $Because" -ForegroundColor Red }
}

$root = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-libraries-" + [guid]::NewGuid().ToString('N'))
try {
    # A drive laid out like the user's F: (2026-11-01)
    foreach ($d in @('DVDs\Jaws', 'Bluray\Heat', 'Documentaries\Cream', 'Basketball', 'Series\Blackadder\Season 1',
                     'Payslips', 'Exeter Uni Essays\Year 1', 'New folder', 'System Volume Information', '$RECYCLE.BIN',
                     '1ac85b24f721833ac835f0', 'Steve Surf Clips\2019')) {
        New-Item -ItemType Directory -Path (Join-Path $root $d) -Force | Out-Null
    }
    Set-Content -LiteralPath (Join-Path $root 'DVDs\Jaws\Jaws-Feature.mp4') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Bluray\Heat\Heat-Feature-BluRay.mkv') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Basketball\Hoop Dreams 1994.mp4') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Series\Blackadder\Season 1\Blackadder-S01-E01.mp4') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Payslips\2024-01.pdf') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Exeter Uni Essays\Year 1\essay.docx') -Value 'x'
    Set-Content -LiteralPath (Join-Path $root 'Steve Surf Clips\2019\wave.MOV') -Value 'x'

    # -----------------------------------------------------------------------
    Write-Host "`nDrive scan and suggestions" -ForegroundColor Cyan
    $c = @(Get-DriveLibraryCandidates -Drive 'X:' -Root $root)
    $names = @($c | ForEach-Object { $_.Name })
    $likely = @($c | Where-Object { $_.Likely } | ForEach-Object { $_.Name })
    Assert-True ('System Volume Information' -notin $names -and '$RECYCLE.BIN' -notin $names) 'system folders are not listed'
    Assert-True ('1ac85b24f721833ac835f0' -notin $names) 'a long hex folder (installer leftovers) is not listed'
    Assert-True ('Payslips' -in $names -and 'Payslips' -notin $likely) 'Payslips is listed but not suggested (no video)'
    Assert-True ('Exeter Uni Essays' -notin $likely -and 'New folder' -notin $likely) 'essays / empty folder not suggested'
    foreach ($n in @('DVDs', 'Bluray', 'Documentaries', 'Series')) { Assert-True ($n -in $likely) "$n suggested" }
    Assert-True ('Basketball' -in $likely) 'Basketball suggested: loose videos in the folder count'
    Assert-True ('Steve Surf Clips' -in $likely) 'a folder with video one level down is suggested (.MOV, any case)'
    Assert-Equal (($names | Sort-Object) -join '|') ($names -join '|') 'listed in name order'
    Assert-Equal 0 @(Get-DriveLibraryCandidates -Drive 'X:' -Root (Join-Path $root 'missing')).Count 'a missing drive gives an empty list'

    # -----------------------------------------------------------------------
    Write-Host "`nWhat a library folder means for the rip" -ForegroundColor Cyan
    Assert-Equal 'movie' (ConvertTo-LibraryChoice 'DVDs').Kind 'DVDs = standard movie folder'
    Assert-Equal 'movie' (ConvertTo-LibraryChoice 'bluray').Kind 'Bluray = standard movie folder (disc type decides)'
    Assert-Equal 'series' (ConvertTo-LibraryChoice 'Series').Kind 'Series = TV series mode'
    Assert-Equal 'Documentary' (ConvertTo-LibraryChoice 'Documentaries').Value 'Documentaries = -Documentary (keeps Part naming)'
    Assert-Equal 'Music' (ConvertTo-LibraryChoice 'Music').Value 'Music = -Music'
    Assert-Equal 'Surf' (ConvertTo-LibraryChoice 'Surf').Value 'Surf = -Surf'
    $g = ConvertTo-LibraryChoice 'Basketball'
    Assert-True ($g.Kind -eq 'genre' -and $g.Value -eq 'Basketball') 'any other folder = -Genre <folder>'
    $g = ConvertTo-LibraryChoice 'TV'
    Assert-True ($g.Kind -eq 'genre' -and $g.Value -eq 'TV') 'a folder called TV is a plain folder, not series mode'

    # -----------------------------------------------------------------------
    Write-Host "`nLibrary prompt" -ForegroundColor Cyan
    $cfg = @('DVDs', 'Documentaries', 'Music', 'Series', 'Basketball')
    # rows: 1 Movies, 2 Documentaries, 3 Music, 4 Series, 5 Basketball (DVDs folded into row 1)
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('')) 6>$null
    Assert-Equal 'movie' $r.Kind 'Enter = standard movie folder (old behaviour)'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @()) 6>$null
    Assert-Equal 'movie' $r.Kind 'end of input = standard movie folder, never an abort'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('2')) 6>$null
    Assert-True ($r.Kind -eq 'switch' -and $r.Value -eq 'Documentary' -and $r.Folder -eq 'Documentaries') 'row 2 = Documentaries'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('4')) 6>$null
    Assert-Equal 'series' $r.Kind 'row 4 = Series'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('9', 'x', '5')) 6>$null
    Assert-True ($r.Kind -eq 'genre' -and $r.Value -eq 'Basketball') 'bad answers re-ask; row 5 = Basketball genre'
    Assert-True ($null -eq (Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('a')) 6>$null)) 'A aborts'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('n', 'Rugby League')) 6>$null
    Assert-True ($r.Kind -eq 'genre' -and $r.Folder -eq 'Rugby League') 'N types a new folder'
    $r = Read-LibraryChoice -Drive 'F:' -Configured $cfg -ReadInput (New-InputQueue @('n', 'a\b', 'n', '', '3')) 6>$null
    Assert-True ($r.Kind -eq 'switch' -and $r.Value -eq 'Music') 'a bad new folder name is refused; blank goes back to the list'

    # Not configured: the drive is scanned; only suggested folders are numbered, [O] shows the rest
    $scan = @(
        [pscustomobject]@{ Name = 'Basketball'; Likely = $true }
        [pscustomobject]@{ Name = 'DVDs'; Likely = $true }
        [pscustomobject]@{ Name = 'Payslips'; Likely = $false }
        [pscustomobject]@{ Name = 'Surf'; Likely = $true })
    # suggested rows: 2 Basketball, 3 Surf
    $r = Read-LibraryChoice -Drive 'F:' -Candidates $scan -ReadInput (New-InputQueue @('3')) 6>$null
    Assert-Equal 'Surf' $r.Value 'not configured: suggested folders are rows 2..'
    $r = Read-LibraryChoice -Drive 'F:' -Candidates $scan -ReadInput (New-InputQueue @('4')) 6>$null
    Assert-True ($null -eq $r -or $r.Kind -eq 'movie') 'an unlisted (hidden) row number is not accepted (input ran out = movie)'
    $r = Read-LibraryChoice -Drive 'F:' -Candidates $scan -ReadInput (New-InputQueue @('o', '3')) 6>$null
    Assert-True ($r.Kind -eq 'genre' -and $r.Value -eq 'Payslips') '[O] lists every folder (row 3 = Payslips after Basketball)'
    $r = Read-LibraryChoice -Drive 'F:' -Candidates @([pscustomobject]@{ Name = 'Stuff'; Likely = $false }) -ReadInput (New-InputQueue @('2')) 6>$null
    Assert-Equal 'Stuff' $r.Value 'nothing suggested: every folder is listed straight away'

    # -----------------------------------------------------------------------
    Write-Host "`nWiring" -ForegroundColor Cyan
    $ripText = Get-Content -LiteralPath $ripDiscPath -Raw
    $setupText = Get-Content -LiteralPath $setupPath -Raw
    $cfgText = Get-Content -LiteralPath $configPath -Raw
    $askIdx = $ripText.IndexOf('Read-LibraryChoice -Drive')
    Assert-True ($askIdx -gt $ripText.IndexOf('Read-OutputDrive -Default') -and $askIdx -lt $ripText.IndexOf('$script:GenreLabel = ')) 'rip-disc.ps1 asks after the output drive and before the genre variables are set'
    Assert-True ($ripText -match "\`$script:LaunchParams\['Genre'\] = \`$Genre" -and $ripText -match "\`$script:LaunchParams\[\`$library\.Value\] = \`$true") 'the choice is carried into the [D] restart'
    Assert-True ($setupText.Contains('libraryFolders = @($libraryFolders)')) 'setup.ps1 saves libraryFolders'
    Assert-True ($setupText.Contains('MediaLibraries.ps1') -and $cfgText.Contains("'MediaLibraries.ps1'")) 'setup.ps1 and Load-Config.ps1 share MediaLibraries.ps1'
    Assert-True ($cfgText.Contains('$script:Config_LibraryFolders')) 'Load-Config.ps1 reads libraryFolders'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
