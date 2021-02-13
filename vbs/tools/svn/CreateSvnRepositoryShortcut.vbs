Option Explicit

' ���|�W�g���ւ̃V���[�g�J�b�g�t�@�C���쐬�X�N���v�g v1.0

'<<�g����>>
'  1. �{�X�N���v�g�Ɠ��K�w�̃t�H���_�Ƀ��|�W�g���p�X���L�ڂ���
'     ���̓t�@�C�����쐬����B
'       [�t�@�C����]
'         <INPUT_PATH_FILE_NAME �L�ڂ̃t�@�C����>.txt
'       [���g�̗�]
'         �P�s�ځFfile:///C:/svn_repo/c/FreeRTOSV7.1.1/Source
'         �Q�s�ځFfile:///C:/svn_repo/c/simple_sio
'  2. �{�X�N���v�g�����s�B
'  
'  �� �{�X�N���v�g�Ɠ��K�w�̃t�H���_�ɃV���[�g�J�b�g�t�@�C�����쐬�����B

Const INPUT_PATH_FILE_NAME = "_repository_path"

Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ExtractTailWord()

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sInputPathFilePath
Dim sCurDirPath
sCurDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
sInputPathFilePath = sCurDirPath & "\" & INPUT_PATH_FILE_NAME & ".txt"

If objFSO.FileExists( sInputPathFilePath ) Then
    'Do Nothing
Else
    MsgBox "�u" & INPUT_PATH_FILE_NAME & ".txt�v�����݂��܂���I"
    MsgBox "�����𒆒f���܂��B"
    WScript.Quit
End If

Dim cRepoFilePaths
Set cRepoFilePaths = CreateObject("System.Collections.ArrayList")

'���|�W�g���t�@�C���p�X�擾
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile( sInputPathFilePath, 1, True ) '�������FIO���[�h�i1:�Ǐo���A2:�V�K�����݁A8:�ǉ������݁j�A��O�����F�V�����t�@�C�����쐬���邩�ǂ���
Do Until objTxtFile.AtEndOfStream
    cRepoFilePaths.Add objTxtFile.ReadLine
Loop
objTxtFile.Close

'�V���[�g�J�b�g�쐬
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sRepoDirPath
For Each sRepoDirPath In cRepoFilePaths
    'MsgBox sRepoDirPath
    Dim sShortcutFilePath
    Dim sShortcutTrgtPath
    Dim sRepoDirName
    sRepoDirName = ExtractTailWord( sRepoDirPath, "/" )
    sShortcutFilePath = sCurDirPath & "\" & sRepoDirName & ".lnk"
    sShortcutTrgtPath = sRepoDirPath
    'MsgBox sRepoDirName & vbNewLine & sShortcutFilePath & vbNewLine & sShortcutTrgtPath
    
    With objWshShell.CreateShortcut( sShortcutFilePath )
        .TargetPath = "TortoiseProc.exe"
        .Arguments = " /command:repobrowser /path:""" & sShortcutTrgtPath & """"
        .Description = sShortcutTrgtPath
        .Save
    End With
Next

MsgBox "���|�W�g���ւ̃V���[�g�J�b�g���쐬���܂����B"

' �O���v���O���� �C���N���[�h�֐�
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

