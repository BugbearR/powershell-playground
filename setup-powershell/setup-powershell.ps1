$latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/PowerShell/PowerShell/releases/latest"
$latestVersion = $latestRelease.tag_name -replace '^v', ''
$currentVersion = $PSVersionTable.PSVersion.ToString()

Write-Host "Current Version: $currentVersion"
Write-Host "Latest Version: $latestVersion"

if ($currentVersion -ne $latestVersion) {
    winget install --id Microsoft.Powershell --source winget
    Write-Host "Please restart shell."
    return 1
}
