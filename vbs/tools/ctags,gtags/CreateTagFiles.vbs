'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�^�O�t�@�C���쐬"

Dim bIsContinue
bIsContinue = True

Dim sTrgtDirPath

If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        sTrgtDirPath = InputBox( "�t�H���_�p�X���w�肵�Ă�������", sPROG_NAME, WScript.Arguments(0) )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sTrgtDirPath = WScript.Env("Current")
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sTrgtDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** �t�H���_���݊m�F ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists( sTrgtDirPath ) Then
        ' Do Nothing
    Else
        MsgBox "�w�肳�ꂽ�t�H���_�����݂��܂���B" & vbNewLine & sTrgtDirPath, vbOKOnly, sPROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
        WScript.Quit
    End If
End If

'*** �^�O�t�@�C���쐬 ***
If bIsContinue = True Then
    MsgBox "�uctags�v�Ɓugtags�v�Ƀp�X���ʂ��Ă��Ȃ��ꍇ�́A�p�X��ʂ��Ă�����s���Ă��������B", vbOk, sPROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "�^�O�t�@�C���̍쐬���������܂����B", vbOk, sPROG_NAME
Else
    'Do Nothing
End If
