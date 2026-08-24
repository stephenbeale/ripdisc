#Requires -Version 5.1
<#
.SYNOPSIS
    Logic tests for the end-of-session log file reminder and its terminal hyperlink.

.DESCRIPTION
    Runs against the real function bodies lifted out of rip-disc.ps1 / continue-rip.ps1
    at run time via the PowerShell AST parser, so the tests cannot drift from the shipped
    code without failing. No disc, MakeMKV or HandBrake needed.

.EXAMPLE
    .\tests\Test-LogFileReminder.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'
$continuePath = Join-Path $repoRoot 'continue-rip.ps1'

foreach ($p in @($ripDiscPath, $continuePath)) {
    if (-not (Test-Path $p)) { throw "Cannot find $p - run this from inside the repo." }
}

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

. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Format-TerminalLink')))
. ([scriptblock]::Create((Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName 'Show-LogFileReminder')))

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

# Format-TerminalLink reads $env:WT_SESSION / $env:TERM_PROGRAM. Swap them for the
# duration of a test and always restore, so the suite cannot leak state into the shell.
function Invoke-WithTerminalEnv {
    param([string]$WtSession, [string]$TermProgram, [scriptblock]$Body)

    $oldWt = $env:WT_SESSION
    $oldTerm = $env:TERM_PROGRAM
    try {
        if ($WtSession) { $env:WT_SESSION = $WtSession } else { Remove-Item Env:\WT_SESSION -ErrorAction SilentlyContinue }
        if ($TermProgram) { $env:TERM_PROGRAM = $TermProgram } else { Remove-Item Env:\TERM_PROGRAM -ErrorAction SilentlyContinue }
        return & $Body
    } finally {
        if ($null -ne $oldWt) { $env:WT_SESSION = $oldWt } else { Remove-Item Env:\WT_SESSION -ErrorAction SilentlyContinue }
        if ($null -ne $oldTerm) { $env:TERM_PROGRAM = $oldTerm } else { Remove-Item Env:\TERM_PROGRAM -ErrorAction SilentlyContinue }
    }
}

$ESC = [char]27

Write-Host "`nTerminal capability detection" -ForegroundColor Cyan

$plain = Invoke-WithTerminalEnv -WtSession $null -TermProgram $null -Body {
    Format-TerminalLink -Uri 'file:///C:/logs/a.log' -Text 'C:\logs\a.log'
}
Assert-Equal 'C:\logs\a.log' $plain 'plain conhost gets the bare path, no escape codes'
Assert-True (-not $plain.Contains($ESC)) 'no ESC character leaks into unsupported terminals'

$wt = Invoke-WithTerminalEnv -WtSession 'abc-123' -TermProgram $null -Body {
    Format-TerminalLink -Uri 'file:///C:/logs/a.log' -Text 'C:\logs\a.log'
}
Assert-Equal "$ESC]8;;file:///C:/logs/a.log$ESC\C:\logs\a.log$ESC]8;;$ESC\" $wt 'Windows Terminal gets a full OSC 8 sequence'

$vscode = Invoke-WithTerminalEnv -WtSession $null -TermProgram 'vscode' -Body {
    Format-TerminalLink -Uri 'file:///C:/logs/a.log' -Text 'C:\logs\a.log'
}
Assert-True ($vscode.Contains($ESC)) 'VS Code terminal gets a hyperlink too'

$other = Invoke-WithTerminalEnv -WtSession $null -TermProgram 'Apple_Terminal' -Body {
    Format-TerminalLink -Uri 'file:///C:/logs/a.log' -Text 'C:\logs\a.log'
}
Assert-Equal 'C:\logs\a.log' $other 'an unrecognised TERM_PROGRAM is treated as unsupported'

Write-Host "`nFile URI construction" -ForegroundColor Cyan
Assert-Equal 'file:///C:/logs/a.log' ([uri]'C:\logs\a.log').AbsoluteUri 'backslashes become a file:// URI'
Assert-Equal 'file:///C:/my%20logs/a%20b.log' ([uri]'C:\my logs\a b.log').AbsoluteUri 'spaces are percent-encoded so the link is clickable'

Write-Host "`nShow-LogFileReminder behaviour" -ForegroundColor Cyan

# Captures Write-Host output by shadowing it inside a child scope.
function Get-ReminderOutput {
    param([string]$LogFile)

    $script:LogFile = $LogFile
    $captured = New-Object System.Collections.Generic.List[string]
    function Write-Host {
        param(
            [Parameter(ValueFromPipeline = $true, Position = 0)]$Object,
            [ConsoleColor]$ForegroundColor
        )
        $captured.Add("$Object")
    }
    Show-LogFileReminder
    return $captured
}

$noLog = Get-ReminderOutput -LogFile $null
Assert-Equal 0 $noLog.Count 'prints nothing when no log file was ever configured'

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ripdisc-log-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
try {
    $missing = Join-Path $tmpDir 'never-written.log'
    $missingOut = (Get-ReminderOutput -LogFile $missing) -join "`n"
    Assert-True ($missingOut -match 'No log file was written') 'says so plainly when the log file does not exist'
    Assert-True ($missingOut -match [regex]::Escape($missing)) 'still shows the expected path when the file is missing'

    $real = Join-Path $tmpDir 'session.log'
    Set-Content -Path $real -Value 'test entry' -Encoding UTF8
    $realOut = (Get-ReminderOutput -LogFile $real) -join "`n"
    Assert-True ($realOut -match 'SESSION LOG') 'prints a clearly labelled section header'
    Assert-True ($realOut -match [regex]::Escape($real)) 'the literal path is always present and copyable'
    Assert-True ($realOut -match [regex]::Escape($tmpDir)) 'the containing folder is shown too'
    Assert-True ($realOut -match 'notepad') 'suggests how to open it'
} finally {
    Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nBoth scripts agree" -ForegroundColor Cyan
$normalise = { param($s) ($s -replace '\s+', ' ').Trim() }
foreach ($fn in @('Format-TerminalLink', 'Show-LogFileReminder')) {
    $a = & $normalise (Import-FunctionFromScript -ScriptPath $ripDiscPath -FunctionName $fn)
    $b = & $normalise (Import-FunctionFromScript -ScriptPath $continuePath -FunctionName $fn)
    # Comments differ deliberately (continue-rip notes it is a copy), so compare code only.
    $stripComments = { param($s) (($s -split "`n" | Where-Object { $_ -notmatch '^\s*#' }) -join ' ') -replace '\s+', ' ' }
    Assert-Equal (& $stripComments $a) (& $stripComments $b) "$fn is identical in both scripts"
}

$total = $script:Passed + $script:Failed
Write-Host "`n$($script:Passed)/$total passed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
exit 0
