# 注意：本スクリプトは管理者権限で実行する必要がある

$args_num = $args.Length
if ($args_num -lt 2) {
	Write-Host "[error] 引数を指定してください。"
	Write-Host "  usage: change_drive_letter.ps1 <label_name> <new_drive_letter>"
	exit
}
$LabelName = $args[0]
$DriveLetterNew = $args[1]
$DriveLetterOld = (Get-volume -friendlyname $LabelName).DriveLetter

if ($DriveLetterNew -eq $DriveLetterOld) {
	Write-Host "[error] ドライブレターはすでに変更済みです。"
} else {
	Get-Partition -DriveLetter $DriveLetterOld | Set-Partition -NewDriveLetter $DriveLetterNew
}

