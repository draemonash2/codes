'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�V���[�g�J�b�g�t�@�C���̎w����t�H���_�Ɉړ�"

Dim bIsContinue
bIsContinue = True

Dim sFilePath

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
       sFilePath = WScript.Arguments(0)
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sFilePath = WScript.Env("Focused")
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sFilePath = "C:\codes\vbs\tools\win\other\test.vbs - �V���[�g�J�b�g.lnk"
    End If
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If bIsContinue = True Then
    If sFilePath = "" Then
        MsgBox "�t�@�C�����I������Ă��܂���", vbYes, sPROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
    If objFSO.GetExtensionName( sFilePath ) <> "lnk" Then
        MsgBox "�V���[�g�J�b�g�t�@�C����I�����Ă�������" & vbNewLine & sFilePath, vbYes, sPROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �V���[�g�J�b�g�w����t�H���_�I�[�v�� ***
Dim objShell
Set objShell = WScript.CreateObject("Shell.Application")
If bIsContinue = True Then
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim sTrgtFilePath
    Dim sTrgtDirPath
    sTrgtFilePath = objWshShell.CreateShortcut( sFilePath ).TargetPath
    sTrgtDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
    If EXECUTION_MODE = 0 Then 'Explorer������s
        objShell.Explore sTrgtDirPath
        Set objShell = Nothing
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        WScript.Open(sFilePath)
    Else '�f�o�b�O���s
        MsgBox "�u" & sTrgtDirPath & "�v���G�N�X�v���[���[�ŊJ���܂�"
        objShell.Explore sTrgtDirPath
        Set objShell = Nothing
    End If
Else
    'Do Nothing
End If
