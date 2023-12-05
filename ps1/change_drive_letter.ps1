# ���ӁF�{�X�N���v�g�͊Ǘ��Ҍ����Ŏ��s����K�v������

$args_num = $args.Length
if ($args_num -lt 2) {
	Write-Host "[error] �������w�肵�Ă��������B"
	Write-Host "  usage: change_drive_letter.ps1 <label_name> <new_drive_letter>"
	exit
}
$LabelName = $args[0]
$DriveLetterNew = $args[1]
$DriveLetterOld = (Get-volume -friendlyname $LabelName).DriveLetter

if ($DriveLetterNew -eq $DriveLetterOld) {
	Write-Host "[error] �h���C�u���^�[�͂��łɕύX�ς݂ł��B"
} else {
	Get-Partition -DriveLetter $DriveLetterOld | Set-Partition -NewDriveLetter $DriveLetterNew
}

