Option Explicit

Dim sExecTime
Dim sBackupBatchFile
If WScript.Arguments.Count = 2 Then
	sBackupBatchFile = WScript.Arguments(0) '���ӁjBackUpFiles.bat.git_sample�̏ꍇ�́A�V���{���b�N�����N���o�R���Ȃ���΃p�X���w�肷�邱�ƁB
	sExecTime = WScript.Arguments(1)
Else
	WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
	WScript.Quit
End If

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Do While True
	Dim sCurTime
	Dim sTrgtDateTime
	sCurTime = Now()
	sTrgtDateTime = Left(sCurTime, InStr(sCurTime, " ")) & " " & sExecTime
	'MsgBox sTrgtDateTime
	Dim lDateDiff
	lDateDiff = DateDiff("n", sCurTime, sTrgtDateTime)
	'MsgBox lDateDiff
	If lDateDiff = 0 Then
		Dim sCmd
		sCmd = "cmd /c """ & sBackupBatchFile & """ ""Scheduled backup."""
		'MsgBox "The time has come!" & vbNewLine & sCmd
		objWshShell.Run sCmd, 0, True
	End If
	WScript.sleep(60000) '60[s]
Loop

