<#
.SYNOPSIS
    Reverses the renames recorded in a rename-manifest.csv written by the -Series flow.

.DESCRIPTION
    rip-disc.ps1 / continue-rip.ps1 write rename-manifest.csv into the DiscN folder
    BEFORE renaming any episode, and copy this script next to it. Running it renames
    every file in the manifest back to its original MakeMKV/HandBrake name.

    Names are resolved relative to the manifest's own folder (not the absolute paths
    recorded in it), so the undo still works after the folder or drive letter changes.
    Extras were moved into the series-level "Specials" folder, so their NewName is
    recorded as ..\..\Specials\<name> (..\Specials\<name> when there is no Season or
    DiscN level); manifests from before that (PR #143) record extras\<name>. Undo moves
    them back into the DiscN folder and removes the Specials / extras folder if that
    leaves it empty. Those shapes are the only paths accepted.

    Safe by default:
      - a file that is missing is skipped with a warning (e.g. a rename that never ran
        because the original run stopped part-way through)
      - a file is never renamed over an existing file - name collisions are skipped
      - rows are undone newest-first
      - -WhatIf shows what would happen without touching anything

.PARAMETER ManifestPath
    Path to rename-manifest.csv. Defaults to the one next to this script.

.EXAMPLE
    & "E:\Series\Silicon Valley\Season 2\Disc1\undo-rename.ps1" -WhatIf

.EXAMPLE
    .\undo-rename.ps1 -ManifestPath "E:\Series\Silicon Valley\Season 2\Disc1\rename-manifest.csv"
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string]$ManifestPath = ""
)

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $PSScriptRoot 'rename-manifest.csv'
}
if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Error "Rename manifest not found: $ManifestPath"
    exit 1
}

$manifestDir = Split-Path -Parent (Resolve-Path -LiteralPath $ManifestPath).ProviderPath
$rows = @(Import-Csv -LiteralPath $ManifestPath)
[array]::Reverse($rows)

$restored = 0
$skipped = 0
$extrasDirs = @{}
foreach ($row in $rows) {
    $newName = "$($row.NewName)"
    $originalName = "$($row.OriginalName)"

    # Bare names only, except that NewName may sit in the series-level Specials folder (one
    # or two levels up) or a PR #143-era extras subfolder - a row pointing anywhere else
    # (other folders, other .. paths, rooted paths) is refused, not followed.
    $newNameOk = ($newName -match '^(?:extras[\\/]|(?:\.\.[\\/]){1,2}Specials[\\/])?[^\\/:]+$') -and ($newName -notmatch '(^|[\\/])\.\.?$')
    if (-not $newName -or -not $originalName -or -not $newNameOk -or $originalName -match '[\\/:]') {
        Write-Warning "Skipping malformed row: '$originalName' / '$newName'"
        $skipped++
        continue
    }

    $currentPath = [System.IO.Path]::GetFullPath((Join-Path $manifestDir $newName))
    $originalPath = Join-Path $manifestDir $originalName
    if ($newName -match '[\\/]') { $extrasDirs[(Split-Path -Parent $currentPath)] = $true }

    if (-not (Test-Path -LiteralPath $currentPath)) {
        if (Test-Path -LiteralPath $originalPath) {
            Write-Warning "Skipping ${newName}: not found, and $originalName already exists (already undone, or never renamed)"
        } else {
            Write-Warning "Skipping ${newName}: file not found"
        }
        $skipped++
        continue
    }
    if (Test-Path -LiteralPath $originalPath) {
        Write-Warning "Skipping ${newName}: $originalName already exists - not overwriting"
        $skipped++
        continue
    }

    if ($PSCmdlet.ShouldProcess($currentPath, "Rename back to $originalName")) {
        try {
            # Move-Item so an extra comes back out of Specials / extras; never overwrites.
            Move-Item -LiteralPath $currentPath -Destination $originalPath -ErrorAction Stop
            Write-Host "  $newName -> $originalName" -ForegroundColor Gray
            $restored++
        } catch {
            Write-Warning "Failed to rename ${newName}: $($_.Exception.Message)"
            $skipped++
        }
    }
}

# Remove the Specials / extras folder if undoing emptied it (it only exists because of
# renames). Specials is shared by every disc, so it stays while any other disc's extras remain.
foreach ($extrasDir in $extrasDirs.Keys) {
    if (-not $WhatIfPreference -and (Test-Path -LiteralPath $extrasDir -PathType Container) -and
        -not (Get-ChildItem -LiteralPath $extrasDir -Force | Select-Object -First 1)) {
        Remove-Item -LiteralPath $extrasDir -ErrorAction SilentlyContinue
    }
}

Write-Host "Undo complete: $restored restored, $skipped skipped$(if ($WhatIfPreference) { ' (WhatIf - nothing changed)' })" -ForegroundColor $(if ($skipped) { 'Yellow' } else { 'Green' })
[pscustomobject]@{ Restored = $restored; Skipped = $skipped }
