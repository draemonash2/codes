Option Explicit

'�I�����@�F�^�X�N�}�l�[�W���[�u�ڍׁv->�uwscript.exe�v�̃^�X�N���I������

Const sSCRIPT_NAME = "����"
Const lMSG_OUT_TIME_S = 10		'10[s]
Const lSLEEP_TIME_MS = 30000	'30000[ms]�i1000[ms] * 30[s]�j= 30[s]
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
			objWshShell.Popup sTime & "�ł�", lMSG_OUT_TIME_S, sTime & "�ł�", vbInformation
			bFinished = True
		Else
			'Do Nothing
		End If
	Else
		bFinished = False
	End If
Loop
