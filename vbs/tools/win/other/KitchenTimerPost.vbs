Option Explicit

Const PROG_NAME = "�L�b�`���^�C�}�["

If Wscript.arguments.count = 0 Then
	MsgBox _
		"���Ԃ��w�肳��Ȃ��������ߏI�����܂�", _
		vbYes, _
		PROG_NAME
Else
	WScript.Sleep Wscript.arguments(0) * 60 * 1000

	MsgBox _
		Wscript.arguments(0) & "�����o�߂��܂���", _
		vbYes, _
		PROG_NAME
End If
