param(
    [Parameter(Mandatory)]
    [string]$SeriesRoot
)

if (!(Test-Path $SeriesRoot)) {
    Write-Error "Series root does not exist: $SeriesRoot"
    return
}

$seriesFolderName = Split-Path $SeriesRoot -Leaf

# Detect season from folder name
if ($seriesFolderName -notmatch 'S(\d{2})') {
    Write-Error "Cannot detect season from folder name '$seriesFolderName'"
    return
}

$seasonNumber = [int]$Matches[1]
$seasonTag    = "S{0:D2}" -f $seasonNumber
$seasonFolder = "Season $seasonNumber"

Write-Host "Season detected: $seasonTag" -ForegroundColor Cyan

# Move to series root safely
Push-Location $SeriesRoot

# Find / rename season folder
$existingSeason = Get-ChildItem -Directory |
    Where-Object { $_.Name -match '^Season\s+\d+' } |
    Select-Object -First 1

if (-not $existingSeason) {
    Write-Error "No Season folder found"
    Pop-Location
    return
}

if ($existingSeason.Name -ne $seasonFolder) {
    Write-Host "Renaming '$($existingSeason.Name)' -> '$seasonFolder'" -ForegroundColor Yellow
    Rename-Item $existingSeason.FullName $seasonFolder
}

$seasonPath = Join-Path $SeriesRoot $seasonFolder
Set-Location $seasonPath

# -------- determine next episode number (scan all discs) --------
$seriesParent = Split-Path $SeriesRoot -Parent

$allSeasonEpisodes =
    Get-ChildItem $seriesParent -Directory |
    Where-Object { $_.Name -match '-S\d{2}\s+D\d{2}$' } |
    ForEach-Object {
        $path = Join-Path $_.FullName $seasonFolder
        if (Test-Path $path) {
            Get-ChildItem $path -File -Filter "*.mp4"
        }
    } |
    Where-Object { $_.Name -match "$seasonTag`E(\d{2})" } |
    ForEach-Object { [int]$Matches[1] }

$nextEpisode = if ($allSeasonEpisodes) {
    ($allSeasonEpisodes | Measure-Object -Maximum).Maximum + 1
} else {
    1
}

Write-Host ("Starting episode number: E{0:D2}" -f [int]$nextEpisode) -ForegroundColor Cyan

# -------- rename files in current folder only --------
Get-ChildItem -File -Filter "*.mp4" |
    # Rename all mp4 files regardless of current name
    Where-Object { $_.Extension -eq ".mp4" }|
    Sort-Object Name |
    ForEach-Object {
        $ep = "{0:D2}" -f $nextEpisode
        $newName = "$seasonTag`E$ep$($_.Extension)"

        Write-Host "  $($_.Name) -> $newName" -ForegroundColor Yellow
        Rename-Item $_ -NewName $newName
        $nextEpisode++
    }

Pop-Location
Write-Host "Series cleanup complete." -ForegroundColor Green
