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

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim asSplitWord
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
'   Call Test_ExtractTailWord()
    Private Sub Test_ExtractTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
        Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
        MsgBox Result
    End Sub
