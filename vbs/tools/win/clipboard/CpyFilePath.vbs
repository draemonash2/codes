'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'�g����
'   CpyFilePath.vbs <file_path1> [<file_path2>]...
'
'���p�I�Ȏg�p���@
'   ���炩���߁uMATCH_DIR_NAME�v�ƁuREMOVE_DIR_LEVEL�v��ݒ肵�Ă���
'   �ꍇ�́A�Y�����镶������������ăR�s�[����B
'   ����ȊO�̏ꍇ�́A���̂܂܃R�s�[����B
'       ex1)
'           set MATCH_DIR_NAME=codes
'           set REMOVE_DIR_LEVEL=1
'           c\codes\aaa\bbb\ccc\test.txt
'               �� bbb\ccc\test.txt ���R�s�[
'       ex2)
'           c\codes\aaa\bbb\ccc\test.txt
'               �� c\codes\aaa\bbb\ccc\test.txt ���R�s�[

'####################################################################
'### �ݒ�
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### �C���N���[�h
'####################################################################
Call Include( "C:\codes\vbs\_lib\String.vbs" )

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�t�@�C���p�X���R�s�["

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    End If
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cFilePaths.Count = 0 Then
        MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ���΃p�X�ɒu�� ***
Dim objWshShellEnv
Set objWshShellEnv = WScript.CreateObject("WScript.Shell").Environment("Process")
Dim sMatchDirName
Dim sRemoveDirLevel
sMatchDirName = objWshShellEnv.Item("MATCH_DIR_NAME")
sRemoveDirLevel = objWshShellEnv.Item("REMOVE_DIR_LEVEL")
Dim oFilePath
If sMatchDirName <> "" And sRemoveDirLevel <> "" Then
    Dim cRltvFilePaths
    Set cRltvFilePaths = CreateObject("System.Collections.ArrayList")
    For Each oFilePath In cFilePaths
        Dim sRltvFilePath
        Call ExtractRelativePath(oFilePath, sMatchDirName, CLng(sRemoveDirLevel), sRltvFilePath)
        'Msgbox oFilePath & "�F" & sRltvFilePath '��debug
        cRltvFilePaths.Add sRltvFilePath
    Next
    Set cFilePaths = cRltvFilePaths
End If

'*** �N���b�v�{�[�h�փR�s�[ ***
If bIsContinue = True Then
    Dim sOutString
    Dim bFirstStore
    bFirstStore = True
    For Each oFilePath In cFilePaths
        If bFirstStore = True Then
            sOutString = oFilePath
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & oFilePath
        End If
    Next
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If

' �O���v���O���� �C���N���[�h�֐�
Private Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sOpenFile = objFSO.GetAbsolutePathName( sOpenFile )
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
