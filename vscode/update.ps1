Write-Host "VSCode ���A�b�v�f�[�g���܂��B`r`n"
$Answer = Read-Host "�{�X�N���v�g���s�O�� VSCode ���I�����Ă��������B`r`n�I���ς݂ł���� ""y"" ����͂��Ă��������B[y/n]"
if ( $Answer -ne "y" ) {
	Read-Host "`r`n�v���O�����𒆒f���܂��B"
	Exit 1
}
Read-Host "`r`n�A�b�v�f�[�g�����s���܂��B"

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

Read-Host "�A�b�v�f�[�g���������܂����B"
