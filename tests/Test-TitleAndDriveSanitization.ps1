#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for Get-SafeTitle (path/filename sanitization) and
    Get-NormalizedDriveLetter (-Drive/-OutputDrive normalization).

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 and
    continue-rip.ps1 at run time via the PowerShell AST parser, so the tests cannot
    drift from the shipped code without failing. No disc, MakeMKV or HandBrake needed.

    Covers two live bugs:
      - A title ending in a dot (e.g. "W." by Oliver Stone) or trailing space produced
        a $safeTitle that didn't match what Windows actually creates on disk (Windows
        silently strips a trailing dot/space from the final path component), causing
        later exact-path comparisons (e.g. the Step 3 working-directory guard) to see a
        mismatch and fail a rip that had actually succeeded.
      - -Drive/-OutputDrive normalization only recognized an already-correct "F:" via
        "-match ':$'" and otherwise appended a colon unconditionally, so "F:\", "F::"
        or a stray trailing space produced a malformed drive letter instead of being
        cleaned up.

.EXAMPLE
    .\tests\Test-TitleAndDriveSanitization.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

# ---------- load the real implementations ----------

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

# ---------- tiny assertion harness ----------

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

# ---------- run the full suite once per script, so both stay in sync ----------

foreach ($scriptPath in @($ripDiscPath, $continuePath)) {
    $scriptName = Split-Path $scriptPath -Leaf

    # Import in an isolated scope per script so Get-SafeTitle/Get-NormalizedDriveLetter
    # from one script's source don't linger and mask a divergence in the other's.
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $scriptPath -FunctionName 'Get-SafeTitle')))
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $scriptPath -FunctionName 'Get-NormalizedDriveLetter')))

    Write-Host "`nGet-SafeTitle - illegal path characters ($scriptName)" -ForegroundColor Cyan
    Assert-Equal 'The Matrix' (Get-SafeTitle 'The Matrix') 'a plain title is untouched'
    Assert-Equal 'Movie_ Subtitle' (Get-SafeTitle 'Movie: Subtitle') 'colon is replaced'
    Assert-Equal 'The Arena Hawaii 05_06 Highlight Reel - ASL' (Get-SafeTitle 'The Arena Hawaii 05/06 Highlight Reel - ASL') 'slash is replaced (regression, PR #131)'
    Assert-Equal 'A_B_C_D_E_F_G_H_' (Get-SafeTitle 'A\B/C:D*E?F"G<H>') 'every illegal character is replaced'

    Write-Host "`nGet-SafeTitle - trailing dot/space (the 'W.' bug) ($scriptName)" -ForegroundColor Cyan
    Assert-Equal 'W' (Get-SafeTitle 'W.') 'a trailing dot is stripped, matching what Windows actually creates on disk'
    Assert-Equal 'W' (Get-SafeTitle 'W. ') 'a trailing space after a trailing dot is also stripped'
    Assert-Equal 'W' (Get-SafeTitle 'W.  . . ') 'a run of trailing dots/spaces in any combination is fully stripped'
    Assert-Equal 'Se7en' (Get-SafeTitle 'Se7en') 'a title with no trailing dot/space is unaffected'
    Assert-Equal 'Mr. Smith' (Get-SafeTitle 'Mr. Smith') 'an internal (non-trailing) dot is left alone'
    Assert-Equal '.hack' (Get-SafeTitle '.hack') 'a leading dot is left alone - only trailing is a Windows creation-time issue'

    Write-Host "`nGet-SafeTitle - degenerate input ($scriptName)" -ForegroundColor Cyan
    $allDots = Get-SafeTitle '...'
    Assert-Equal $false ([string]::IsNullOrWhiteSpace($allDots)) 'a title that is nothing but dots still yields a non-empty safe title (fallback)'
    Assert-Equal '...' $allDots 'the fallback for an all-illegal/dots title is the untrimmed sanitized form'

    Write-Host "`nGet-NormalizedDriveLetter ($scriptName)" -ForegroundColor Cyan
    Assert-Equal 'F:' (Get-NormalizedDriveLetter 'F:') 'an already-correct drive letter is unchanged'
    Assert-Equal 'F:' (Get-NormalizedDriveLetter 'F') 'a bare letter gets exactly one trailing colon'
    Assert-Equal 'F:' (Get-NormalizedDriveLetter 'F:\') 'a trailing backslash is stripped, not left before the colon'
    Assert-Equal 'F:' (Get-NormalizedDriveLetter 'F::') 'a doubled colon collapses to exactly one'
    Assert-Equal 'F:' (Get-NormalizedDriveLetter ' F: ') 'surrounding whitespace is trimmed'
    Assert-Equal 'F:' (Get-NormalizedDriveLetter ' F ') 'whitespace and a missing colon are both fixed at once'
    Assert-Equal 'f:' (Get-NormalizedDriveLetter 'f') 'case is preserved (drive letters are case-insensitive on Windows)'
}

# ---------- both scripts must agree ----------
# Get-SafeTitle's comment legitimately differs by one clause between the two scripts
# (rip-disc.ps1 mentions "TMDb lookups" as a display-only use; continue-rip.ps1 has no
# TMDb lookups to mention) - matching the pre-existing convention already used for the
# equivalent hand-written comment both functions were extracted from. Behavioural
# equivalence is what's tested (the same battery of inputs above ran against both
# scripts' copy and produced identical results); only the textually-identical
# Get-NormalizedDriveLetter is checked for exact source equality here.

Write-Host "`nBoth scripts agree" -ForegroundColor Cyan
Assert-Equal (Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Get-NormalizedDriveLetter') (Import-FunctionFromScript -ScriptPath $continuePath -FunctionName 'Get-NormalizedDriveLetter') 'Get-NormalizedDriveLetter is identical in both scripts'

# ---------- summary ----------

Write-Host ""
Write-Host "$script:Passed/$($script:Passed + $script:Failed) passed" -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
if ($script:Failed -gt 0) { exit 1 }
