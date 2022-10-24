Option Explicit

Const sTRGT_TIME = "09:01"
Const sCMD = "cmd /c C:\codes\BackUpFiles.bat"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Do While True
	Dim sCurTime
	Dim sTrgtDateTime
	sCurTime = Now()
	sTrgtDateTime = Left(sCurTime, InStr(sCurTime, " ")) & " " & sTRGT_TIME
	'MsgBox sTrgtDateTime
	Dim lDateDiff
	lDateDiff = DateDiff("n", sCurTime, sTrgtDateTime)
	'MsgBox lDateDiff
	If lDateDiff = 0 Then
		objWshShell.Run sCMD, 0, True
		'MsgBox "Time is Money!"
	End If
	WScript.sleep(60000) '60[s]
Loop


