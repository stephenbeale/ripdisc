#Requires -Version 5.1
<#
.SYNOPSIS
    Tests for the summary tables ("Discovered Metadata", "Ready to rip") in rip-disc.ps1.

.DESCRIPTION
    Runs the real Split-TableText / Format-InfoTable / Write-InfoTable bodies lifted out of
    rip-disc.ps1 via the PowerShell AST parser.

.EXAMPLE
    .\tests\Test-InfoTable.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$ripDiscPath = Join-Path $repoRoot 'rip-disc.ps1'

$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ripDiscPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) { throw "$ripDiscPath has $($parseErrors.Count) parse error(s)." }
foreach ($name in @('Split-TableText', 'Format-InfoTable', 'Write-InfoTable')) {
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true) | Select-Object -First 1
    if (-not $fn) { throw "Function '$name' not found" }
    . ([scriptblock]::Create($fn.Extent.Text))
}

$script:Passed = 0
$script:Failed = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$Because)
    if ($Expected -ceq $Actual) { $script:Passed++; Write-Host "  PASS  $Because" -ForegroundColor Green }
    else {
        $script:Failed++
        Write-Host "  FAIL  $Because" -ForegroundColor Red
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}
function Row { param($Label, $Value, $Color = $null) [pscustomobject]@{ Label = $Label; Value = $Value; Color = $Color } }

Write-Host "`nSplit-TableText" -ForegroundColor Cyan
Assert-Equal 'short' ((Split-TableText -Text 'short' -Width 10) -join '|') 'short text is one piece'
Assert-Equal 'with Don|Cheadle' ((Split-TableText -Text 'with Don Cheadle' -Width 10) -join '|') 'breaks at a space'
Assert-Equal 'abcdefghij|klm' ((Split-TableText -Text 'abcdefghijklm' -Width 10) -join '|') 'no space: cut hard'

Write-Host "`nFormat-InfoTable" -ForegroundColor Cyan
$t = @(Format-InfoTable -Heading 'Discovered Metadata' -Rows @((Row 'Title' 'Miles Ahead'), (Row 'Year' '2015'), (Row 'Season' $null), (Row 'IMDb' '')))
$expected = @(
    '+------------------------------+',
    '| Discovered Metadata          |',
    '+-------+----------------------+',
    '| Title | Miles Ahead          |',
    '| Year  | 2015                 |',
    '+-------+----------------------+')
Assert-Equal ($expected -join "`n") (($t | ForEach-Object { $_.Text }) -join "`n") 'two-column table; empty values left out; value column at least 20 wide'
Assert-Equal 1 @($t | ForEach-Object { $_.Text.Length } | Select-Object -Unique).Count 'every line the same width'

$t = @(Format-InfoTable -Heading 'Ready to rip: Miles Ahead' -Rows @((Row 'Output folder' 'G:\DVDs\Miles Ahead (2015)'), (Row 'Using' 'Drive F:' 'Yellow')))
Assert-Equal '| Output folder | G:\DVDs\Miles Ahead (2015) |' $t[3].Text 'label column fits the longest label'
Assert-Equal 'Yellow' $t[4].Color 'row color carried'
Assert-Equal 'White' $t[3].Color 'default color White'

$t = @(Format-InfoTable -Heading 'H' -MaxWidth 40 -Rows @((Row 'About' 'with Don Cheadle, Emayatzy Corinealdi and a much longer cast list here')))
$rows = @($t | Where-Object { $_.Kind -eq 'row' })
Assert-Equal $true ($rows.Count -gt 1) 'long value wraps onto extra lines'
Assert-Equal $true (@($t | Where-Object { $_.Text.Length -gt 40 }).Count -eq 0) 'no line wider than -MaxWidth'
Assert-Equal '|       |' ($rows[1].Text.Substring(0, 9)) 'wrapped lines leave the label blank'

$t = @(Format-InfoTable -Heading 'Ready to rip: A Very Long Film Title That Goes On And On' -Rows @((Row 'Using' 'Drive F:')))
Assert-Equal 1 @($t | ForEach-Object { $_.Text.Length } | Select-Object -Unique).Count 'a long heading widens the table, all lines still equal'

Write-Host "`nWrite-InfoTable" -ForegroundColor Cyan
$ok = $true
try { Write-InfoTable -Heading 'Smoke' -Rows @((Row 'Title' 'Miles Ahead'), (Row 'Using' 'Drive F:' 'Yellow')) } catch { $ok = $false; Write-Host $_ -ForegroundColor Red }
Assert-Equal $true $ok 'prints without error'

$ripText = [IO.File]::ReadAllText($ripDiscPath)
Assert-Equal $true ($ripText.Contains('Write-InfoTable -Heading "Discovered Metadata"')) 'Discovered Metadata is a table'
Assert-Equal $true ($ripText.Contains('Write-InfoTable -Heading "Ready to rip: $title"')) 'Ready to rip is a table'

Write-Host "`n$script:Passed passed, $script:Failed failed" -ForegroundColor $(if ($script:Failed) { 'Red' } else { 'Green' })
if ($script:Failed) { exit 1 }
