Option Explicit

'<<�T�v>>
'  �w�肵���t�@�C�����o�b�N�A�b�v����B
'  
'<<�g�p���@>>
'  BackUpFile.vbs <filepath> [<backupnum>] [<logfilepath>]
'  
'<<�d�l>>
'  �E�t�@�C�����w�肷��ƌ��ݎ�����t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'  �E�����t�@�C�����̂��̂����݂��Ă�����A�A���t�@�x�b�g��t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'     ex. 211201a, 211202b, �c
'  �E�������Ɏw�肳�ꂽ�o�b�N�A�b�v�����t�@�C�������܂�����A�Â����̂���폜����B
'  �E���s���ʂ͑�O�����Ɏw�肳�ꂽ���O�t�@�C���ɏo�͂���B
'  �E�O��o�b�N�A�b�v������X�V����Ă��Ȃ��ꍇ�A�o�b�N�A�b�v���Ȃ��B
'  
'<<���ӎ���>>
'  �E�o�b�N�A�b�v�Ώۂ̓t�@�C���̂݁B
'  �E�ȉ���S�Ė������ꍇ�A�V�����t�@�C�����X�V����Ă������ߗv���ӁB
'      - �o�b�N�A�b�v�t�@�C���̐ڔ�����"z"�ƂȂ��Ă���t�@�C�������� (ex. file.txt.#b#211122z.txt)
'  �E�ȉ��̗��R�ōŐV/�ŌÃo�b�N�A�b�v�t�@�C������ɍX�V������p���Ȃ��B�����܂�
'    �o�b�N�A�b�v�������������t�@�C�����Ŕ��f����B
'      ����ČÂ��o�b�N�A�b�v�t�@�C�����X�V���Ă��܂����ꍇ�A�t�@�C�������
'      ���t���Â��̂ɍX�V�������V�����t�@�C�����ł��Ă��܂��B
'      �X�V���������Ƃɔ��肷��ƁA��L�̃t�@�C�����폜���ꂸ�A�c���Ă��܂����߁B

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()
                                                            'CreateDirectry()
                                                            'GetFileInfo()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const bEXEC_TEST = False '�e�X�g�p
Const sSCRIPT_NAME = "�t�@�C���o�b�N�A�b�v"
Const sBAK_DIR_NAME = "_#b#"
Const sBAK_FILE_SUFFIX = "#b#"
Const lBAK_FILE_NUM_DEFAULT = 30

'===============================================================================
'= �{����
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= ���C���֐�
'===============================================================================
Public Sub Main()
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim sBakSrcFilePath
    Dim lBakFileNumMax
    Dim sBakLogFilePath
    If cArgs.Count >= 3 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = CLng(cArgs(1))
        sBakLogFilePath = cArgs(2)
    ElseIf cArgs.Count >= 2 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = CLng(cArgs(1))
        sBakLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & objFSO.GetBaseName(WScript.ScriptName) & ".log"
    ElseIf cArgs.Count >= 1 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = lBAK_FILE_NUM_DEFAULT
        sBakLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & objFSO.GetBaseName(WScript.ScriptName) & ".log"
    Else
        WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
        Exit Sub
    End If
    'MsgBox sBakSrcFilePath & vbNewLine & lBakFileNumMax & vbNewLine & sBakLogFilePath
    
    Dim objLogFile
    Set objLogFile = objFSO.OpenTextFile(sBakLogFilePath, 8, True) 'AddWrite
    
    '****************
    '*** ���O���� ***
    '****************
    '�Ώۃt�@�C�����擾
    Dim sBakSrcParDirPath
    Dim sBakSrcFileName
    Dim sBakSrcFileBaseName
    Dim sBakSrcFileExt
    Dim sDateSuffix
    sBakSrcParDirPath = objFSO.GetParentFolderName( sBakSrcFilePath )
    sBakSrcFileName = objFSO.GetFileName( sBakSrcFilePath )
    sBakSrcFileBaseName = objFSO.GetBaseName( sBakSrcFilePath )
    sBakSrcFileExt = objFSO.GetExtensionName( sBakSrcFilePath )
    sDateSuffix = ConvDate2String(Now(),2)
    'MsgBox sBakSrcParDirPath & vbNewLine & sBakSrcFileName & vbNewLine & sBakSrcFileBaseName & vbNewLine & sBakSrcFileExt & vbNewLine & sDateSuffix
    
    '�g���q�L���`�F�b�N
    Dim bExistsExt
    If ( (sBakSrcFileBaseName <> "") And (sBakSrcFileExt <> "") ) Then
        bExistsExt = True
    Else
        bExistsExt = False
    End If
    
    '�o�b�N�A�b�v�t�@�C�����쐬
    Dim sBakDstDirPath
    Dim sBakDstPathBase
    sBakDstDirPath = sBakSrcParDirPath & "\" & sBAK_DIR_NAME
    sBakDstPathBase = sBakDstDirPath & "\" & sBakSrcFileName & "." & sBAK_FILE_SUFFIX
    
    '****************************
    '*** �t�@�C���o�b�N�A�b�v ***
    '****************************
    '�o�b�N�A�b�v�t�H���_�쐬
    Call CreateDirectry( sBakDstDirPath )
    
    '�t�@�C���ꗗ�擾
    Dim cFileList
    Set cFileList = CreateObject("System.Collections.ArrayList")
    Call GetFileListCmdClct( sBakDstDirPath, cFileList, 1, "*")
    
    '�����̍ŐV�t�@�C���T��
    Dim sBakDstFilePathLatest  '�����̍ŐV�o�b�N�A�b�v�t�@�C��
    sBakDstFilePathLatest = ""
    Dim sFilePath
    For Each sFilePath In cFileList
        If ( InStr(sFilePath, sBakDstPathBase) > 0 ) Then
            sBakDstFilePathLatest = sFilePath
        End If
    Next
    Set cFileList = Nothing
    
    '�o�b�N�A�b�v�t�@�C�����m��
    Dim sBakDstFilePath
    '�����̃o�b�N�A�b�v�t�@�C�������݂��A�������t�̃o�b�N�A�b�v�t�@�C�������݂���ꍇ
    If sBakDstFilePathLatest <> "" And _
       InStr(sBakDstFilePathLatest, sBakDstPathBase & sDateSuffix) > 0 Then
        Dim sTailChar
        If bExistsExt = True Then
            sTailChar = Right( objFSO.GetBaseName( sBakDstFilePathLatest ), 1)
        Else
            sTailChar = Right( sBakDstFilePathLatest, 1)
        End If
        Dim lBakDstAlphaIdx
        If Asc(sTailChar) >= Asc("a") And Asc(sTailChar) < Asc("z") Then
            lBakDstAlphaIdx = Asc(sTailChar) + 1
        ElseIf Asc(sTailChar) = Asc("z") Then
            lBakDstAlphaIdx = Asc(sTailChar)
        ElseIf Asc(sTailChar) >= Asc("0") And Asc(sTailChar) <= Asc("9") Then
            lBakDstAlphaIdx = Asc("a")
        Else
            objLogFile.WriteLine "�s���ȃo�b�N�A�b�v�t�@�C����������܂����B"
            objLogFile.WriteLine "  " & sBakDstFilePathLatest
            objLogFile.WriteLine "�v���O�����𒆒f���܂��B"
            Exit Sub
        End If
        If bExistsExt = True Then
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx) & "." & sBakSrcFileExt
        Else
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx)
        End If
    Else
        If bExistsExt = True Then
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & "." & sBakSrcFileExt
        Else
            sBakDstFilePath = sBakDstPathBase & sDateSuffix
        End If
    End If
    'objLogFile.WriteLine sBakDstFilePath & " : " & sBakDstFilePathLatest
    'WScript.Echo sBakDstFilePath & " : " & sBakDstFilePathLatest
    'Exit Sub
    
    '�X�V�����擾
    Dim vDateLastModifiedLatestBk
    Dim vDateLastModifiedTrgt
    Dim bRet
    bRet = GetFileInfo( sBakDstFilePathLatest, 11, vDateLastModifiedLatestBk)
    bRet = GetFileInfo( sBakSrcFilePath, 11, vDateLastModifiedTrgt)
    
    '�����̃o�b�N�A�b�v�t�@�C�������� or �X�V����Ă���ꍇ
    Dim sLogMsg
    sLogMsg = ""
    If ( sBakDstFilePathLatest = "" ) Or _
       ( ( sBakDstFilePathLatest <> "" ) And ( vDateLastModifiedTrgt > vDateLastModifiedLatestBk ) ) Then
        '�t�@�C���o�b�N�A�b�v
        objFSO.CopyFile sBakSrcFilePath, sBakDstFilePath, True
        sLogMsg = "[Success] " & sBakSrcFilePath & " -> " & sBakDstFilePath & "."
    Else
        '�O��o�b�N�A�b�v������X�V����Ă��Ȃ��ꍇ�A�o�b�N�A�b�v���������𒆒f����
        objLogFile.WriteLine "[Skip]    " & sBakSrcFilePath & "."
        Exit Sub
    End If
    
    '************************
    '*** �Â��t�@�C���폜 ***
    '************************
    '�t�@�C�����X�g�擾
    Dim cFileListAll
    Set cFileListAll = CreateObject("System.Collections.ArrayList")
    Call GetFileListCmdClct( sBakDstDirPath, cFileListAll, 1, "*")
    Set cFileList = CreateObject("System.Collections.ArrayList")
    For Each sFilePath in cFileListAll
        If InStr(sFilePath, sBakDstPathBase) > 0 Then
            cFileList.Add sFilePath
        End If
    Next
    
    '�o�b�N�A�b�v�t�@�C���폜
    Dim lBakFileNum
    Dim lDelFileNum
    lBakFileNum = cFileList.Count
    lDelFileNum = 0
    For Each sFilePath In cFileList
        If lBakFileNum > lBakFileNumMax Then
            'objFSO.DeleteFile sFilePath, True
            Call MoveToTrushBox(objFSO, sFilePath)
            lDelFileNum = lDelFileNum + 1
        End If
        lBakFileNum = lBakFileNum - 1
    Next
    Set cFileList = Nothing
    
    If lDelFileNum > 0 Then
        objLogFile.WriteLine sLogMsg & " " & lDelFileNum & " files deleted."
    Else
        objLogFile.WriteLine sLogMsg
    End If
    
    'objLogFile.WriteLine "�o�b�N�A�b�v�����I", vbOKOnly, sSCRIPT_NAME
    
    objLogFile.Close
End Sub

'===============================================================================
'= �����֐�
'===============================================================================
Private Function MoveToTrushBox(ByRef objFSO, ByVal sTrgtPath)
    If objFSO.FileExists(sTrgtPath) Then
        CreateObject("Shell.Application").Namespace(10).movehere sTrgtPath
        Do While objFSO.FileExists(sTrgtPath) Or objFSO.FolderExists(sTrgtPath)
            '�폜�����͔񓯊��Ői�s���邽�߁A�폜���ɃX�N���v�g���I������ƍ폜�����͒��f�����B
            '�폜�Ώۂ��폜�����܂őҋ@����B
            WScript.sleep(100)
        Loop
        MoveToTrushBox = True
    Else
        MoveToTrushBox = False
    End If
End Function

'===============================================================================
'= �e�X�g�֐�
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim sDesktopPath
    sDesktopPath = objWshShell.SpecialFolders("Desktop")
    
    MsgBox "=== test start ==="
    
    Dim sTrgtFilePath
    Dim sTrgtFilePathOrg
    Dim sBakDirPath
    Dim sBakLogName
    Dim objTxtFile
    Select Case lTestCase
        Case 1
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            sBakLogName = sDesktopPath & "\backup_test.log"
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            cArgs.Add 5
            cArgs.Add sBakLogName
            
            Call Main()
            MsgBox "1 �o�b�N�A�b�v������(����ǉ�)"
            
            Dim objDummyFile
            Set objDummyFile = objFSO.OpenTextFile(sDesktopPath & "\" & sBAK_DIR_NAME & "\dummy_file.txt", 8, True)
            objDummyFile.WriteLine "a"
            objDummyFile.Close
            
            Call Main()
            MsgBox "2 �o�b�N�A�b�v������(�ω��Ȃ�)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "3 �o�b�N�A�b�v������(a�ǉ�)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "4 �o�b�N�A�b�v������(b�ǉ�)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "5 �o�b�N�A�b�v������(c�ǉ�)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "6 �o�b�N�A�b�v������(d�ǉ�)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "7 �o�b�N�A�b�v������(e�ǉ� ����폜)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "8 �o�b�N�A�b�v������(f�ǉ� a�폜)"
            
            cArgs(1) = 2
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "9 �o�b�N�A�b�v������(g�ǉ� b,c,d,e�폜)"
        Case 2
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            cArgs.Add 5
            Call Main()
        Case 3
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            Call Main()
        Case 4
            Dim sTrgtFilePath1
            Dim sTrgtFilePath2
            Dim sTrgtFilePath3
            Dim sTrgtFilePath4
            sTrgtFilePath1 = sDesktopPath & "\backup_test.txt"
            sTrgtFilePath2 = sDesktopPath & "\.backup_test.txt"
            sTrgtFilePath3 = sDesktopPath & "\backup_test"
            sTrgtFilePath4 = sDesktopPath & "\.backup_test"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            sBakLogName = sDesktopPath & "\backup_test.log"
            
            If objFSO.FileExists(sTrgtFilePath1) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath1, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath2) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath2, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath3) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath3, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath4) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath4, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath1
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath2
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath3
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath4
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
        Case Else
            Call Main()
    End Select
    
    MsgBox "=== test finished ==="
End Sub '}}}

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}
