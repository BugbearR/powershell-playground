# Japanese 106/109 Keyboard Setup Script
# This script must be run with administrator privileges

# Check administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run with administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as administrator and try again." -ForegroundColor Red
    exit 1
}

Write-Host "Starting Japanese 106/109 keyboard setup..." -ForegroundColor Green

# Function to set registry value safely
function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value,
        [string]$Type = "String"
    )
    
    try {
        if (!(Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
            Write-Host "Created registry path: $Path" -ForegroundColor Yellow
        }
        
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
        Write-Host "Set $Path\$Name = $Value" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Failed to set $Path\$Name : $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# 1. Set keyboard layout to Japanese 106/109
Write-Host "`nConfiguring keyboard layout to Japanese 106/109..." -ForegroundColor Yellow

# Registry paths for keyboard configuration
$keyboardLayoutPath = "HKLM:\SYSTEM\CurrentControlSet\Services\i8042prt\Parameters"
$keyboardClassPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout"
$keyboardMapPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\00000411"

# Set keyboard subtype to Japanese 106/109
# SubType values: 0 = AT Enhanced, 2 = Japanese 106/109, 3 = Korean 101, etc.
$success1 = Set-RegistryValue -Path $keyboardLayoutPath -Name "LayerDriver JPN" -Value "kbd106.dll"
$success2 = Set-RegistryValue -Path $keyboardLayoutPath -Name "OverrideKeyboardIdentifier" -Value "PCAT_106KEY"
$success3 = Set-RegistryValue -Path $keyboardLayoutPath -Name "OverrideKeyboardSubtype" -Value "2" -Type "DWord"
$success4 = Set-RegistryValue -Path $keyboardLayoutPath -Name "OverrideKeyboardType" -Value "7" -Type "DWord"

# Set keyboard layout DLL for Japanese
$success5 = Set-RegistryValue -Path $keyboardClassPath -Name "DosKeyboardLayout" -Value "00000411" -Type "DWord"

# Ensure Japanese keyboard layout is registered
if (Test-Path $keyboardMapPath) {
    $success6 = Set-RegistryValue -Path $keyboardMapPath -Name "Layout File" -Value "KBDJPN.DLL"
    $success7 = Set-RegistryValue -Path $keyboardMapPath -Name "Layout Text" -Value "Japanese"
} else {
    Write-Host "Japanese keyboard layout registry path not found. Creating..." -ForegroundColor Yellow
    New-Item -Path $keyboardMapPath -Force | Out-Null
    $success6 = Set-RegistryValue -Path $keyboardMapPath -Name "Layout File" -Value "KBDJPN.DLL"
    $success7 = Set-RegistryValue -Path $keyboardMapPath -Name "Layout Text" -Value "Japanese"
}

# 2. Configure input method for current user
Write-Host "`nConfiguring input method for current user..." -ForegroundColor Yellow

try {
    # Set Japanese as input language
    $inputLanguageList = New-WinUserLanguageList -Language ja-JP
    $inputLanguageList[0].InputMethodTips.Clear()
    $inputLanguageList[0].InputMethodTips.Add('0411:00000411')  # Japanese IME
    Set-WinUserLanguageList -LanguageList $inputLanguageList -Force
    Write-Host "Japanese input method configured successfully." -ForegroundColor Green
} catch {
    Write-Host "Error configuring input method: $($_.Exception.Message)" -ForegroundColor Red
}

# 3. Set keyboard hardware configuration
Write-Host "`nConfiguring keyboard hardware settings..." -ForegroundColor Yellow

# Additional hardware-specific settings
$hardwarePath = "HKLM:\SYSTEM\CurrentControlSet\Services\i8042prt\Parameters"
$success8 = Set-RegistryValue -Path $hardwarePath -Name "KeyboardDataQueueSize" -Value "100" -Type "DWord"
$success9 = Set-RegistryValue -Path $hardwarePath -Name "NumberOfButtons" -Value "3" -Type "DWord"

# Set scan code mapping for Japanese layout
$scanCodePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout"
# Remove any existing scan code maps that might interfere
try {
    Remove-ItemProperty -Path $scanCodePath -Name "Scancode Map" -ErrorAction SilentlyContinue
    Write-Host "Cleared existing scancode mappings." -ForegroundColor Green
} catch {
    Write-Host "No existing scancode mappings to clear." -ForegroundColor Yellow
}

# 4. Configure for all users (default user profile)
Write-Host "`nConfiguring keyboard settings for all new users..." -ForegroundColor Yellow

# Load default user hive
try {
    $defaultUserPath = "C:\Users\Default\NTUSER.DAT"
    if (Test-Path $defaultUserPath) {
        reg load "HKU\DefaultUser" "$defaultUserPath" 2>$null
        
        # Set keyboard layout for default user
        $defaultUserKeyboardPath = "HKU:\DefaultUser\Keyboard Layout\Preload"
        if (!(Test-Path $defaultUserKeyboardPath)) {
            New-Item -Path $defaultUserKeyboardPath -Force | Out-Null
        }
        Set-ItemProperty -Path $defaultUserKeyboardPath -Name "1" -Value "00000411"
        
        # Unload default user hive
        [gc]::Collect()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Host "Default user keyboard configuration completed." -ForegroundColor Green
    }
} catch {
    Write-Host "Warning: Could not configure default user settings: $($_.Exception.Message)" -ForegroundColor Yellow
}

# 5. Install Japanese IME if not present
Write-Host "`nChecking Japanese IME installation..." -ForegroundColor Yellow

try {
    $japaneseIME = Get-WinUserLanguageList | Where-Object { $_.LanguageTag -eq "ja-JP" }
    if ($japaneseIME) {
        Write-Host "Japanese IME is already installed." -ForegroundColor Green
    } else {
        Write-Host "Installing Japanese IME..." -ForegroundColor Yellow
        $languageList = Get-WinUserLanguageList
        $languageList.Add("ja-JP")
        Set-WinUserLanguageList -LanguageList $languageList -Force
        Write-Host "Japanese IME installation completed." -ForegroundColor Green
    }
} catch {
    Write-Host "Error checking/installing Japanese IME: $($_.Exception.Message)" -ForegroundColor Red
}

# 6. Verify keyboard configuration
Write-Host "`nVerifying keyboard configuration..." -ForegroundColor Cyan

try {
    $currentLayout = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\i8042prt\Parameters" -Name "OverrideKeyboardSubtype" -ErrorAction SilentlyContinue
    $layoutDLL = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\i8042prt\Parameters" -Name "LayerDriver JPN" -ErrorAction SilentlyContinue
    
    if ($currentLayout -and $currentLayout."OverrideKeyboardSubtype" -eq 2) {
        Write-Host "Keyboard subtype: Japanese 106/109 (Value: 2)" -ForegroundColor Green
    } else {
        Write-Host "Keyboard subtype: Not set or incorrect" -ForegroundColor Red
    }
    
    if ($layoutDLL -and $layoutDLL."LayerDriver JPN" -eq "kbd106.dll") {
        Write-Host "Keyboard driver: kbd106.dll (Japanese 106/109)" -ForegroundColor Green
    } else {
        Write-Host "Keyboard driver: Not set or incorrect" -ForegroundColor Red
    }
    
    $userLanguages = Get-WinUserLanguageList
    $japaneseFound = $userLanguages | Where-Object { $_.LanguageTag -eq "ja-JP" }
    if ($japaneseFound) {
        Write-Host "Japanese input method: Configured" -ForegroundColor Green
    } else {
        Write-Host "Japanese input method: Not configured" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error during verification: $($_.Exception.Message)" -ForegroundColor Red
}

# Summary of success/failure
Write-Host "`nConfiguration Summary:" -ForegroundColor Cyan
$totalSuccess = ($success1 -and $success2 -and $success3 -and $success4 -and $success5 -and $success6 -and $success7 -and $success8 -and $success9)

if ($totalSuccess) {
    Write-Host "All keyboard configuration steps completed successfully!" -ForegroundColor Green
} else {
    Write-Host "Some configuration steps failed. Please check the output above." -ForegroundColor Yellow
}

# Completion message
Write-Host "`nJapanese 106/109 keyboard setup completed!" -ForegroundColor Green
Write-Host "Please restart the system to apply all changes." -ForegroundColor Yellow
Write-Host "`nAfter restart, you may need to:" -ForegroundColor Cyan
Write-Host "1. Go to Settings > Time & Language > Language & region" -ForegroundColor White
Write-Host "2. Click on Japanese language and select 'Options'" -ForegroundColor White
Write-Host "3. Verify that Microsoft IME is installed and configured" -ForegroundColor White
Write-Host "4. Use Alt+Shift or Windows+Space to switch between input methods" -ForegroundColor White

# Restart confirmation
$restart = Read-Host "`nWould you like to restart now? (Y/N)"
if ($restart -eq "Y" -or $restart -eq "y") {
    Write-Host "Restarting the system..." -ForegroundColor Yellow
    Restart-Computer -Force
} else {
    Write-Host "Please restart manually later to apply all changes." -ForegroundColor Yellow
}
