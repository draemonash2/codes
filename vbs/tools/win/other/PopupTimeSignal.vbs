Option Explicit

Const lIntervalSecond = 180 '[s] = 60[s] * 3[min]
Const lWaitMSecond = 60000 '[ms] = 1000[ms] * 60[s]
Const lMsgOutputSecond = 5 '[s]

Do While 1
	WScript.sleep(lWaitMSecond) '1•ª‘Ò‚¿
	If Fix(Timer()) Mod ( lIntervalSecond ) = 0 Then
		Dim sTime
		sTime = Time
		sTime = Left(sTime, InStrRev(sTime, ":") - 1)
		Dim objWshShell
		Set objWshShell = WScript.CreateObject("WScript.Shell")
		objWshShell.Popup sTime & "‚Å‚·", lMsgOutputSecond, sTime & "‚Å‚·", vbInformation
	End If
Loop
