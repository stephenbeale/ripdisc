#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for the MakeMKV progress display helpers (--progress=-same parsing, duration
    formatting and ETA).

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 at run time via the
    PowerShell AST parser, so the tests cannot drift from the shipped code without failing.

.EXAMPLE
    .\tests\Test-MakeMkvProgress.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

if (-not (Test-Path $ripDiscPath)) { throw "Cannot find $ripDiscPath - run this from inside the repo." }

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

foreach ($fnName in 'ConvertFrom-MakeMkvProgressLine', 'Format-RipDuration', 'Get-RipEta') {
    . ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $fnName)))
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

Write-Host "`nConvertFrom-MakeMkvProgressLine" -ForegroundColor Cyan

$p = ConvertFrom-MakeMkvProgressLine -Line "Current progress - 42%  , Total progress - 17% "
Assert-Equal 42 $p.Current "parses the current-title percentage"
Assert-Equal 17 $p.Total "parses the total percentage"
$p = ConvertFrom-MakeMkvProgressLine -Line "Current progress - 100%, Total progress - 100%"
Assert-Equal 100 $p.Total "tolerates tighter spacing and 100%"
Assert-Equal $null (ConvertFrom-MakeMkvProgressLine -Line "Current operation: Saving to MKV file") "returns null for an operation line"
Assert-Equal $null (ConvertFrom-MakeMkvProgressLine -Line "Title #1 was added (7 cell(s), 0:28:27)") "returns null for ordinary MakeMKV output"
Assert-Equal $null (ConvertFrom-MakeMkvProgressLine -Line "Error 'Scsi error' occurred while reading 'X' at offset '123'") "returns null for an error line"

Write-Host "`nFormat-RipDuration" -ForegroundColor Cyan

Assert-Equal "45s" (Format-RipDuration ([TimeSpan]::FromSeconds(45))) "seconds only under a minute"
Assert-Equal "12m 03s" (Format-RipDuration ([TimeSpan]::FromSeconds(723))) "minutes with zero-padded seconds"
Assert-Equal "1h 05m" (Format-RipDuration ([TimeSpan]::FromMinutes(65))) "hours with zero-padded minutes"
Assert-Equal "0s" (Format-RipDuration ([TimeSpan]::Zero)) "zero duration"

Write-Host "`nGet-RipEta" -ForegroundColor Cyan

Assert-Equal $null (Get-RipEta -Elapsed ([TimeSpan]::FromMinutes(1)) -TotalPercent 0) "no ETA at 0% (would divide by zero)"
Assert-Equal $null (Get-RipEta -Elapsed ([TimeSpan]::FromMinutes(30)) -TotalPercent 100) "no ETA once complete"
Assert-Equal 1800 ([int](Get-RipEta -Elapsed ([TimeSpan]::FromMinutes(10)) -TotalPercent 25).TotalSeconds) "10 min for 25% leaves 30 min"
Assert-Equal 600 ([int](Get-RipEta -Elapsed ([TimeSpan]::FromMinutes(10)) -TotalPercent 50).TotalSeconds) "10 min for 50% leaves 10 min"

Write-Host "`nWatchdog regression guard" -ForegroundColor Cyan

$ripDiscText = Get-Content $ripDiscPath -Raw
Assert-True ($ripDiscText -match '--progress=-same') "MakeMKV is launched with --progress=-same"
Assert-True ($ripDiscText -notmatch 'Saving \d\+ titles\|Current progress') "progress lines no longer mark the rip as started (they also appear during the disc scan)"

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
