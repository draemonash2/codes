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
'	1.1.0	2020/02/09	sleep��
'==============================================================================

'==============================================================================
' �ݒ�
'==============================================================================
Const PROG_NAME = "�L�b�`���^�C�}�["

Dim lWaitMinites
lWaitMinites = InputBox( "�҂�����(��)����͂��Ă�������", PROG_NAME, 1 )

If lWaitMinites = 0 Then
	MsgBox "�L�����Z�����܂���", vbYes, PROG_NAME
Else
	Dim vAnswer
	vAnswer = MsgBox( lWaitMinites & "���Ԃ̃^�C�}�[��ݒ肵�܂���", vbOkCancel, PROG_NAME )
	If vAnswer <> vbOk Then
		MsgBox "�L�����Z�����܂���", vbYes, PROG_NAME
	Else
		WScript.sleep(lWaitMinites * 60 * 1000)
		MsgBox lWaitMinites & "�����o�߂��܂���", vbYes, lWaitMinites & "���o��"
	End If
End If

