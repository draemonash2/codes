'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s
'
' �g����
'   CpyFilePath.vbs [-d <match_dir_name>] [-l <rem_dir_level>] <file_path1> [<file_path2>]...
'   
'     ex1)
'         CpyFilePath.vbs c\test1.txt c\test2.txt
'             �� c\test1.txt<���s>c\test2.txt ���R�s�[
'     ex2)
'         CpyFilePath.vbs -d codes -l 1 c\codes\aaa\bbb\ccc\test.txt c\codes\aaa\bbb\a.txt
'             �� bbb\ccc\test.txt<���s>bbb\a.txt ���R�s�[
'     
'     (��) EXECUTION_MODE = 1 �ɂđ��΃p�X�u���������ꍇ�́A�{�X�N���v�g���s�O��
'          sMatchDirNamesRemoveDirLevel�ɒl��ݒ肵�Ă�������
'          �iex. sMatchDirName = "codes" sRemoveDirLevel = "1"�j

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
Dim sMatchDirName
Dim sRemoveDirLevel

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim bIsGetDirNameTiming
        Dim bIsGetDirLevelTiming
        bIsGetDirNameTiming = False
        bIsGetDirLevelTiming = False
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            Select Case sArg
                Case "-d"
                    bIsGetDirNameTiming = True
                    bIsGetDirLevelTiming = False
                Case "-l"
                    bIsGetDirLevelTiming = True
                    bIsGetDirNameTiming = False
                Case Else
                    If bIsGetDirNameTiming = True Then
                        sMatchDirName = sArg
                        bIsGetDirNameTiming = False
                    ElseIf bIsGetDirLevelTiming = True Then
                        sRemoveDirLevel = sArg
                        bIsGetDirLevelTiming = False
                    Else
                        cFilePaths.add sArg
                    End If
            End Select
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
