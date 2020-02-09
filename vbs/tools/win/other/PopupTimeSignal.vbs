Option Explicit

'終了方法：タスクマネージャー「詳細」->「wscript.exe」のタスクを終了する

Const sSCRIPT_NAME = "時報"
Const lMSG_OUT_TIME_S = 10		'10[s]
Const lSLEEP_TIME_MS = 30000	'30000[ms]（1000[ms] * 30[s]）= 30[s]
Const lINTERVEL_TIME_MIN = 30	'30[min]

Dim bFinished
bFinished = False

Do While 1
	WScript.sleep(lSLEEP_TIME_MS)
	If (Fix(Timer/60) mod lINTERVEL_TIME_MIN) = 0 Then
		If bFinished = False Then
			Dim sTime
			sTime = Time
			sTime = Left(sTime, InStrRev(sTime, ":") - 1)
			Dim objWshShell
			Set objWshShell = WScript.CreateObject("WScript.Shell")
			objWshShell.Popup sTime & "です", lMSG_OUT_TIME_S, sTime & "です", vbInformation
			bFinished = True
		Else
			'Do Nothing
		End If
	Else
		bFinished = False
	End If
Loop
