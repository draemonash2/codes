Write-Host "VSCode をアップデートします。`r`n"
$Answer = Read-Host "本スクリプト実行前に VSCode を終了してください。`r`n終了済みであれば ""y"" を入力してください。[y/n]"
if ( $Answer -ne "y" ) {
	Read-Host "`r`nプログラムを中断します。"
	Exit 1
}
Read-Host "`r`nアップデートを実行します。"

# Remove temp file from portable user data
Remove-Item -Recurse -Force -Path "data/user-data" -Include @("Backups", "Cache", "CachedData", "GPUCache", "logs")

# Download latest stable build
curl.exe -L "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-archive" -o stable.zip

# Delete anything except user data, update script and downloaded zip file
Get-ChildItem -Exclude @("data", "_update.ps1", "_update.ps1.lnk", "stable.zip") | Remove-Item -Recurse -Force

# Unzip it
Expand-Archive -Path "stable.zip" -DestinationPath .

# Delete downloaded package
Remove-Item -Path "stable.zip"

Read-Host "アップデートが成功しました。"
