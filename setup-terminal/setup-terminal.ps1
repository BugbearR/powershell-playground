# Script to set PowerShell 7 as the default shell for Windows Terminal
# Recommended to run with administrator privileges

param(
    [switch]$Force,
    [switch]$Backup
)

# Search for PowerShell 7 executable paths
$pwsh7Paths = @(
    "C:\Program Files\PowerShell\7\pwsh.exe",
    "C:\Users\$env:USERNAME\AppData\Local\Microsoft\powershell\pwsh.exe"
)

$pwsh7Path = $null
foreach ($path in $pwsh7Paths) {
    if (Test-Path $path) {
        $pwsh7Path = $path
        break
    }
}

# Handle case when PowerShell 7 is not found
if (-not $pwsh7Path) {
    Write-Error "PowerShell 7 not found. Please install PowerShell 7."
    Write-Host "Download PowerShell 7: https://github.com/PowerShell/PowerShell/releases"
    exit 1
}

Write-Host "PowerShell 7 found: $pwsh7Path" -ForegroundColor Green

# Windows Terminal settings file path
$settingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

# Check if Windows Terminal is installed
if (-not (Test-Path $settingsPath)) {
    Write-Error "Windows Terminal settings file not found. Please verify Windows Terminal is installed."
    exit 1
}

Write-Host "Windows Terminal settings file: $settingsPath" -ForegroundColor Green

# Create backup
if ($Backup) {
    $backupPath = "$settingsPath.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Copy-Item $settingsPath $backupPath
    Write-Host "Settings file backed up to: $backupPath" -ForegroundColor Yellow
}

try {
    # Load settings file
    $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
    
    # Search for or create PowerShell 7 profile
    $pwsh7Profile = $null
    $pwsh7Guid = "{574e775e-4f2a-5b96-ac1e-a2962a402336}"  # Standard GUID for PowerShell 7
    
    # Search for PowerShell 7 in existing profiles
    foreach ($profile in $settings.profiles.list) {
        if ($profile.commandline -like "*pwsh.exe*" -or $profile.source -eq "Windows.Terminal.PowershellCore") {
            $pwsh7Profile = $profile
            $pwsh7Guid = $profile.guid
            break
        }
    }
    
    # Create PowerShell 7 profile if it doesn't exist
    if (-not $pwsh7Profile) {
        Write-Host "Creating PowerShell 7 profile..." -ForegroundColor Yellow
        
        $newProfile = @{
            guid = $pwsh7Guid
            name = "PowerShell 7"
            commandline = $pwsh7Path
            icon = "ms-appx:///ProfileIcons/pwsh.png"
            startingDirectory = "%USERPROFILE%"
            hidden = $false
        }
        
        $settings.profiles.list += $newProfile
        Write-Host "PowerShell 7 profile added" -ForegroundColor Green
    } else {
        Write-Host "Using existing PowerShell 7 profile: $($pwsh7Profile.name)" -ForegroundColor Green
        # Update command line path
        $pwsh7Profile.commandline = $pwsh7Path
    }
    
    # Set default profile
    $currentDefault = $settings.defaultProfile
    $settings.defaultProfile = $pwsh7Guid
    
    # Write back to settings file
    $settings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
    
    Write-Host "`nConfiguration completed!" -ForegroundColor Green
    Write-Host "Default profile changed: $currentDefault -> $pwsh7Guid" -ForegroundColor Cyan
    Write-Host "PowerShell 7 will launch the next time you open Windows Terminal." -ForegroundColor Green
    
} catch {
    Write-Error "Error occurred while updating settings: $($_.Exception.Message)"
    exit 1
}

# Check if Windows Terminal process is running and prompt for restart
$wtProcess = Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue
if ($wtProcess) {
    Write-Host "`nNote: Windows Terminal is currently running." -ForegroundColor Yellow
    Write-Host "Please restart Windows Terminal to apply the settings." -ForegroundColor Yellow
    
    if ($Force) {
        Write-Host "Force closing Windows Terminal..." -ForegroundColor Red
        $wtProcess | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "Please restart Windows Terminal." -ForegroundColor Green
    }
}

Write-Host "`nScript execution completed!" -ForegroundColor Green
