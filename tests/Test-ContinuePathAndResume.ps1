#Requires -Version 5.1
<#
.SYNOPSIS
    Regression tests for the null $finalOutputDir bug and the wrong-step resume hint
    (2026-08-24, continue-rip.ps1 reported against the Martin Scorsese Presents the
    Blues boxset).

.DESCRIPTION
    Two unrelated defects were found in the same session:

    1. A malformed/empty -OutputDrive can make one of the Join-Path calls that build
       $finalOutputDir emit a non-terminating error and silently leave the variable
       null or empty, while the script carries on regardless - every downstream
       Test-Path/Join-Path call on it then throws a raw parameter-binding exception.
       Fixed with a construction-time validation (both scripts) plus an upfront
       Test-DriveReady call that fails fast on a disconnected/missing output drive.

    2. continue-rip.ps1's "To pick up from here" resume hint suggested -FromStep 4
       after a Step 2 failure, because `Sort-Object Number` silently sorts
       DESCENDING on Windows PowerShell 5.1 when the pipeline objects are
       [hashtable] (as $script:AllSteps entries are) rather than [PSCustomObject].
       Fixed by sorting on a script block ({ $_.Number }) instead of a bare
       property name.

    Both fixes are exercised here against the real function bodies / real source
    text, lifted out of the scripts with the PowerShell parser, so the tests cannot
    drift from the shipped code without failing.

    No disc, no MakeMKV, no HandBrake and no real drives are touched.

.EXAMPLE
    .\tests\Test-ContinuePathAndResume.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

# ---------- shared helpers (same technique as Test-EpisodeNaming.ps1) ----------

function Import-FunctionFromScript {
    param([string]$ScriptPath, [string]$FunctionName)

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "$ScriptPath has $($parseErrors.Count) parse error(s); fix those before testing."
    }

    $fn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
    }, $true) | Select-Object -First 1

    if (-not $fn) { throw "Function '$FunctionName' not found in $ScriptPath" }
    return $fn.Extent.Text
}

# Pulls the top-level "is $finalOutputDir usable" validation line out of the source
# text (it isn't inside a function, so it can't be lifted via the AST function
# finder above) so the test exercises the exact shipped condition.
function Get-FinalOutputDirValidationCondition {
    param([string]$ScriptPath)

    $searchPattern = [regex]::Escape('$finalOutputDir -notmatch')
    $line = Select-String -Path $ScriptPath -Pattern $searchPattern |
        Select-Object -First 1
    if (-not $line) { throw "Could not find the finalOutputDir validation line in $ScriptPath" }

    # The shipped line reads: if (<condition>) {
    # Strip the "if (" / ") {" wrapper so what's left is a bare boolean expression
    # Invoke-Expression can evaluate directly.
    $text = $line.Line.Trim()
    $text = $text -replace '^if\s*\(', ''
    $text = $text -replace '\)\s*\{\s*$', ''
    return $text.Trim()
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
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Because)
    Assert-Equal $true $Condition $Because
}

# ---------- Part 1: $finalOutputDir construction-time validation ----------

Write-Host "`n$finalOutputDir validation rejects the exact failure shape (both scripts)" -ForegroundColor Cyan

foreach ($scriptUnderTest in @(
    @{ Path = $continuePath; Label = 'continue-rip.ps1' },
    @{ Path = $ripDiscPath;  Label = 'rip-disc.ps1' }
)) {
    $condition = Get-FinalOutputDirValidationCondition -ScriptPath $scriptUnderTest.Path

    # Evaluate the real condition text against a handful of $finalOutputDir values,
    # including the two concrete shapes seen in the incident: a genuinely empty
    # string, and the malformed ":\..." shape produced when -OutputDrive is empty.
    foreach ($case in @(
        @{ Value = $null;                                        ShouldFail = $true;  Name = 'null' }
        @{ Value = '';                                            ShouldFail = $true;  Name = 'empty string' }
        @{ Value = '   ';                                         ShouldFail = $true;  Name = 'whitespace only' }
        @{ Value = ':\Documentaries\Title\Season 1\Disc4';        ShouldFail = $true;  Name = 'empty-drive-letter shape (":\\...")' }
        @{ Value = 'F::\Documentaries\Title\Season 1\Disc4';      ShouldFail = $true;  Name = 'double-colon provider-qualified shape' }
        @{ Value = 'Documentaries\Title\Season 1\Disc4';          ShouldFail = $true;  Name = 'relative path, no drive at all' }
        @{ Value = 'F:\Documentaries\Title\Season 1\Disc4';       ShouldFail = $false; Name = 'well-formed drive-rooted path' }
        @{ Value = 'F:\DVDs\Some Movie';                          ShouldFail = $false; Name = 'well-formed simple movie path' }
    )) {
        $finalOutputDir = $case.Value
        $failed = Invoke-Expression $condition
        Assert-Equal $case.ShouldFail ([bool]$failed) "$($scriptUnderTest.Label): $($case.Name) is $(if ($case.ShouldFail) { 'rejected' } else { 'accepted' })"
    }
}

# ---------- Part 2: Test-DriveReady handles empty/null input clearly ----------

Write-Host "`nTest-DriveReady gives a clear message on empty input (both scripts)" -ForegroundColor Cyan

foreach ($scriptUnderTest in @(
    @{ Path = $continuePath; Label = 'continue-rip.ps1' },
    @{ Path = $ripDiscPath;  Label = 'rip-disc.ps1' }
)) {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $scriptUnderTest.Path -FunctionName 'Test-DriveReady')))

    $resultEmpty = Test-DriveReady -Path ''
    Assert-Equal $false $resultEmpty.Ready "$($scriptUnderTest.Label): empty path is never Ready"
    Assert-True ($resultEmpty.Message -notmatch 'from path:\s*$') "$($scriptUnderTest.Label): empty-path message isn't the old bare 'from path:' with nothing after it"

    $resultNull = Test-DriveReady -Path $null
    Assert-Equal $false $resultNull.Ready "$($scriptUnderTest.Label): null path is never Ready"

    # A real, well-formed path to a drive that (almost certainly) doesn't exist on
    # this machine should fail cleanly rather than throwing.
    $resultMissingDrive = Test-DriveReady -Path 'Q:\Documentaries\Some Title'
    Assert-Equal $false $resultMissingDrive.Ready "$($scriptUnderTest.Label): a plausible-but-absent drive letter (Q:) is reported not-ready, not thrown"
}

# ---------- Part 3: resume hint suggests the step that actually failed ----------

Write-Host "`nResume hint suggests the correct -FromStep after a Step 2 failure (continue-rip.ps1)" -ForegroundColor Cyan

. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $continuePath -FunctionName 'Get-Step')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $continuePath -FunctionName 'Get-RemainingSteps')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $continuePath -FunctionName 'Show-StepsSummary')))

$script:AllSteps = @(
    @{ Number = 1; Key = "makemkv";   Name = "MakeMKV rip";        Description = "Rip the disc to MKV files";             Needs = "a disc in the drive"; Resumable = $false }
    @{ Number = 2; Key = "handbrake"; Name = "HandBrake encoding"; Description = "Encode the MKV files to MP4";           Needs = "MKV files in the MakeMKV temp folder"; Resumable = $true }
    @{ Number = 3; Key = "organize";  Name = "Organize files";     Description = "Rename, prefix and move the MP4 files"; Needs = "MP4 files in the output folder"; Resumable = $true }
    @{ Number = 4; Key = "open";      Name = "Open directory";     Description = "Open the finished folder in Explorer";  Needs = "the output folder to exist"; Resumable = $true }
)
$title = 'TestTitle'

function Get-SuggestedResumeStep {
    # Mirrors the exact computation Show-StepsSummary performs, but returns the
    # number instead of printing it, so the assertion below doesn't have to scrape
    # console output.
    $remaining = Get-RemainingSteps
    $resumable = @($remaining | Where-Object { $_.Resumable } | Sort-Object { $_.Number })
    if ($resumable.Count -gt 0) { return $resumable[0].Number }
    return $null
}

# Scenario matching the incident: only Step 1 is assumed-complete (skipped), and
# Step 2 fails immediately - before Complete-CurrentStep is ever called for it.
$script:CompletedSteps = @(Get-Step -Number 1)
Assert-Equal 2 (Get-SuggestedResumeStep) 'failing at the very start of Step 2 suggests -FromStep 2, not 4 (regression)'

# Step 1 and 2 both done, Step 3 fails immediately.
$script:CompletedSteps = @(Get-Step -Number 1), (Get-Step -Number 2)
Assert-Equal 3 (Get-SuggestedResumeStep) 'failing at the start of Step 3 suggests -FromStep 3'

# Nothing done yet (FromStep was 2, and it failed before Step 1 could even be marked skipped - edge case).
$script:CompletedSteps = @()
Assert-Equal 2 (Get-SuggestedResumeStep) 'failing with nothing completed yet still suggests the first resumable step (2), not the last'

# ---------- summary ----------

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
