Option Explicit

'==============================================================================
'�y�����z
'	�w�肵������(��)�̌o�߂�҂��āA���b�Z�[�W��\������
'
'�y�g�p���@�z
'	1) KitchenTimer.vbs �����s���āA�҂����Ԃ���͂���
'
'�y�o�������z
'	�Ȃ�
'
'�y���������z
'	1.0.0	2019/08/03	�V�K�쐬
'	1.1.0	2019/09/26	�����N���Ή�
'==============================================================================

'==============================================================================
' �ݒ�
'==============================================================================
Const PROG_NAME = "�L�b�`���^�C�}�["

Dim lWaitMinites
lWaitMinites = InputBox( "�҂�����(��)����͂��Ă�������", PROG_NAME, 1 )

If lWaitMinites = 0 Then
	MsgBox _
		"�L�����Z�����܂���", _
		vbYes, _
		PROG_NAME
Else
	Dim vAnswer
	vAnswer = MsgBox( _
		lWaitMinites & "���Ԃ̃^�C�}�[��ݒ肵�܂���", _
		vbOkCancel, _
		PROG_NAME _
	)
	If vAnswer <> vbOk Then
		MsgBox _
			"�L�����Z�����܂���", _
			vbYes, _
			PROG_NAME
	Else
		Dim objWsh
		Set objWsh = WScript.CreateObject("WScript.Shell")
		objWsh.Run "KitchenTimerPost.vbs "& lWaitMinites
	End If
End If

