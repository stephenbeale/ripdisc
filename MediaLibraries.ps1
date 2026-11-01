<#
.SYNOPSIS
    Finds media server library folders (Jellyfin/Plex) on an output drive: F:\DVDs,
    F:\Documentaries, F:\Music ... Dot-sourced by Load-Config.ps1 (the rip-time library
    prompt) and setup.ps1 (step 5, which saves them as libraryFolders). No side effects.
#>
# Folders that are never libraries: Windows/system folders, chkdsk's FOUND.nnn, recycle bin.
function Test-SystemFolderName {
    param([string]$Name)
    return ($Name -match '^(\$|System Volume Information$|FOUND\.\d{3}$|Recovery$|Config\.Msi$|MSOCache$|PerfLogs$|Windows$|Program Files|ProgramData$|Users$|Documents and Settings$|OneDriveTemp$|\.)' -or
            $Name -match '^[0-9a-f]{16,}$')
}

# Is $Path a folder holding video, loose or in title folders ("Documentaries\Cream\
# Cream-Feature.mp4"), or named like a library? Looks at most $MaxChildren sub-folders, two levels deep, so a big
# library on a slow USB drive is still quick. Returns $true / $false.
function Test-LikelyLibraryFolder {
    param([string]$Path, [int]$MaxChildren = 12)
    $name = Split-Path $Path -Leaf
    if (Test-SystemFolderName $name) { return $false }
    if ($name -match '(?i)^(movies?|films?|dvds?|blu-?rays?|bd|4k|uhd|documentar(y|ies)|docs|series|tv|tv ?shows?|shows|anime|kids|cartoons|music ?videos?|concerts?|music|surf|sports?|tutorials?|fitness|exercise|stand-?up|comedy)$') { return $true }
    $videoExt = '^\.(mp4|mkv|m4v|avi|mov|wmv|ts|m2ts|iso)$'
    try {
        # Videos straight in the folder count too (F:\Basketball holds loose .mp4 files)
        if (@([System.IO.Directory]::EnumerateFiles($Path) | Where-Object { [System.IO.Path]::GetExtension($_) -match $videoExt } | Select-Object -First 1).Count -gt 0) { return $true }
        $children = @([System.IO.Directory]::EnumerateDirectories($Path) | Select-Object -First $MaxChildren)
        foreach ($c in $children) {
            if (@([System.IO.Directory]::EnumerateFiles($c) | Where-Object { [System.IO.Path]::GetExtension($_) -match $videoExt } | Select-Object -First 1).Count -gt 0) { return $true }
            foreach ($g in @([System.IO.Directory]::EnumerateDirectories($c) | Select-Object -First 4)) {
                if (@([System.IO.Directory]::EnumerateFiles($g) | Where-Object { [System.IO.Path]::GetExtension($_) -match $videoExt } | Select-Object -First 1).Count -gt 0) { return $true }
            }
        }
    } catch { }
    return $false
}

# Top-level folders of $Drive ("F:"), each @{ Name; Likely }, system folders left out,
# sorted by name. Tests pass -Root to a temp folder.
function Get-DriveLibraryCandidates {
    param([string]$Drive, [string]$Root = "")
    if (-not $Root) { $Root = "$("$Drive".TrimEnd('\'))\" }
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { -not (Test-SystemFolderName $_.Name) -and -not ($_.Attributes -band [IO.FileAttributes]::System) } |
        Sort-Object Name | ForEach-Object {
            [pscustomobject]@{ Name = $_.Name; Likely = [bool](Test-LikelyLibraryFolder $_.FullName) }
        })
}
