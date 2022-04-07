Option Explicit

'==============================================================================
'�y�����z
'	�w�肵������(��)�̌o�߂�҂��āA���b�Z�[�W��\������
'
'�y�g�p���@�z
'	1) KitchenTimer.vbs �����s���āA�҂�����(��)����͂���
'
'�y�o�������z
'	�Ȃ�
'
'�y���������z
'	1.0.0	2019/08/03	�V�K�쐬
'	1.1.0	2019/09/26	�����N���Ή�
'	1.1.1	2020/02/09	sleep��
'	1.1.2	2020/08/21	�b/���ԕ\���Ή�
'	1.2.0	2021/01/14	�^�C�g���o�͋@�\�ǉ�
'	1.2.1	2021/01/30	�f�t�H���g�^�C�g���C��
'	1.2.2	2022/04/07	���s���t�@�C���o��/�폜�����ǉ�
'==============================================================================

'==============================================================================
'= �C���N���[�h
'==============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ConvDate2String()

'==============================================================================
' �ݒ�
'==============================================================================
Const sPROG_NAME = "�L�b�`���^�C�}�["

'==============================================================================
'= �{����
'==============================================================================
'*************************************************
'* �^�C�}�[�ݒ�
'*************************************************
Dim dWaitMinites
dWaitMinites = InputBox( "�҂�����(��)����͂��Ă�������", sPROG_NAME, 1 )
If IsEmpty(dWaitMinites) = True Then
	MsgBox "�L�����Z�����܂���", vbYes, sPROG_NAME
	WScript.Quit
ElseIf dWaitMinites = 0 Then
	MsgBox "�L�����Z�����܂���", vbYes, sPROG_NAME
	WScript.Quit
End If

Dim sOutputMsg
sOutputMsg = InputBox( "�^�C�g������͂��Ă�������", sPROG_NAME )
If IsEmpty(sOutputMsg) = True Then
	MsgBox "�L�����Z�����܂���", vbYes, sPROG_NAME
	WScript.Quit
End If

Dim dWaitTime
Dim sWaitTimeUnit
If dWaitMinites < 1 Then
	dWaitTime = Round( dWaitMinites * 60, 2 )
	sWaitTimeUnit = "�b"
ElseIf dWaitMinites >= 60 Then
	dWaitTime = Round( dWaitMinites / 60, 2 )
	sWaitTimeUnit = "����"
Else
	dWaitTime = Round( dWaitMinites, 2 )
	sWaitTimeUnit = "��"
End IF

Dim vAnswer
vAnswer = MsgBox( dWaitTime & sWaitTimeUnit & "�̃^�C�}�[��ݒ肵�܂���", vbOkCancel, sPROG_NAME )
If IsEmpty(vAnswer) = True Then
	MsgBox "�L�����Z�����܂���", vbYes, sPROG_NAME
	WScript.Quit
ElseIf vAnswer <> vbOk Then
	MsgBox "�L�����Z�����܂���", vbYes, sPROG_NAME
	WScript.Quit
End If

'*************************************************
'* ���s���t�@�C���o��
'*************************************************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sStartTime
sStartTime = ConvDate2String(Now(), 0)
Dim sTrgtFilePath
sTrgtFilePath = objWshShell.SpecialFolders("Desktop") & "\running_kitchentimer_" & dWaitMinites & "min[" & sOutputMsg & "]_" & sStartTime & ".txt"
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objTxtFile
On Error Resume Next
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
objTxtFile.WriteLine "running..."
objTxtFile.Close
On Error Goto 0

'*************************************************
'* �҂�����
'*************************************************
WScript.sleep(dWaitMinites * 60 * 1000) 'x[min] * 60[s] * 1000[ms]

'*************************************************
'* ���b�Z�[�W�o��
'*************************************************
If sOutputMsg = "" Then
	MsgBox sPROG_NAME & vbNewLine & dWaitTime & sWaitTimeUnit & "���o�߂��܂���", vbYes, dWaitTime & sWaitTimeUnit & "�o��"
Else
	MsgBox sPROG_NAME & vbNewLine & dWaitTime & sWaitTimeUnit & "���o�߂��܂���", vbYes, sOutputMsg
End If

'*************************************************
'* ���s���t�@�C���폜
'*************************************************
On Error Resume Next
objFSO.DeleteFile sTrgtFilePath, True
On Error Goto 0

'==============================================================================
'= �C���N���[�h�֐�
'==============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
