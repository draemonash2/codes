Option Explicit

Const PROG_NAME = "キッチンタイマー"

If Wscript.arguments.count = 0 Then
	MsgBox _
		"時間が指定されなかったため終了します", _
		vbYes, _
		PROG_NAME
Else
	WScript.Sleep Wscript.arguments(0) * 60 * 1000

	MsgBox _
		Wscript.arguments(0) & "分が経過しました", _
		vbYes, _
		PROG_NAME
End If
