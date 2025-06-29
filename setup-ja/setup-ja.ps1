# https://qiita.com/bibou6/items/0a136bca349050d42b20

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run with administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as administrator and try again." -ForegroundColor Red
    exit 1
}

#日本語の言語パックインストール
Install-Language ja-JP -CopyToSettings

#表示言語を日本語に変更する
Set-systemPreferredUILanguage ja-JP

#UIの言語を日本語で上書きします
Set-WinUILanguageOverride -Language ja-JP

#時刻/日付の形式をWindowsの言語と同じにします
Set-WinCultureFromLanguageListOptOut -OptOut $False

#タイムゾーンを東京にします
Set-TimeZone -id "Tokyo Standard Time"

#ロケーションを日本にします
Set-WinHomeLocation -GeoId 0x7A

#システムロケールを日本にします
Set-WinSystemLocale -SystemLocale ja-JP

#ユーザーが使用する言語を日本語にします
Set-WinUserLanguageList -LanguageList ja-JP,en-US -Force

#入力する言語を日本語で上書きします
Set-WinDefaultInputMethodOverride -InputTip "0411:00000411"

#ようこそ画面と新しいユーザーの設定をコピーします
Copy-UserInternationalSettingsToSystem -welcomescreen $true -newuser $true
