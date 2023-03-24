'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )          ' ShowFolderSelectDialog()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )              ' ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\SettingFileClass.vbs" )    ' SettingFile

'===============================================================================
'= �ݒ�l
'===============================================================================
Const bEXEC_TEST = False '�e�X�g�p
Const sPROG_NAME = "�Q�ƃt�@�C������"
Const lADD_DATE_TYPE = 1 '�t�^��������̎�ʁi1:���ݓ����A2:�t�@�C��/�t�H���_�X�V�����j
Const lDATE_STR_TYPE = 1
Const bEVACUATE_ORG_FILE = True
Const bCHOOSE_DOWNLOAD_DIR_PATH = False
Const bCHOOSE_FILE_AT_DIALOG_BOX = True
Const sSHORTCUT_FILE_SUFFIX = "s"
Const sORIGINAL_FILE_PREFIX = "o"
Const sEDIT_FILE_PREFIX     = "e"
Const sTEMP_FILE_NAME = "CopyRefFile.cfg"

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
    Dim sSrcParDirPath
    Dim sIniDstParDirPath
    Dim cSelectedPaths
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    
    '*** �I���t�@�C���擾 ***
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim sArg
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In cArgs
            cSelectedPaths.add sArg
            If sSrcParDirPath = "" Then
                sSrcParDirPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sSrcParDirPath = WScript.Env("Current")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sSrcParDirPath = "C:\Users\draem\OneDrive\�f�X�N�g�b�v"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\codes\vbs\tools\win\file_ope\CpyAndAddModDate.vbs"
        cSelectedPaths.Add "C:\codes\vbs\tools\win\file_ope\CpyAndAddNowDate.vbs"
        cSelectedPaths.Add "C:\codes\vbs\tools\win\other"
    End If
    
    '*** �t�@�C���p�X�`�F�b�N ***
    If cSelectedPaths.Count = 0 Then
        MsgBox "�t�@�C��/�t�H���_���I������Ă��܂���B", vbOKOnly, sPROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    
    '*** �㏑���m�F ***
    '���s���x�����߂邽�߁A�㏑���m�F�ȗ�
    'Dim vbAnswer
    'vbAnswer = MsgBox( "���Ƀt�@�C�������݂��Ă���ꍇ�A�㏑������܂��B���s���܂����H", vbOkCancel, sPROG_NAME )
    'If vbAnswer = vbOk Then
    '    'Do Nothing
    'Else
    '    MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
    '    Exit Sub
    'End If
    
    '*** �o�͐�I�� ***
    '�o�͐�t�H���_�p�X�擾 from �N���b�v�{�[�h
    'sIniDstParDirPath = CreateObject("htmlfile").ParentWindow.Clipboarddata.GetData("text")
    'If objFSO.FolderExists( sIniDstParDirPath ) = False Then
    '    sIniDstParDirPath = objWshShell.SpecialFolders("Desktop")
    'End If
    
    '�o�͐�t�H���_�p�X�擾 from �ݒ�t�@�C��
    Dim clSetting
    Set clSetting = New SettingFile
    Dim sSettingFilePath
    sSettingFilePath = objFSO.GetSpecialFolder(2) & "\" & sTEMP_FILE_NAME
    
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sIniDstParDirPath, objWshShell.SpecialFolders("Desktop"), False)
    
    Dim sDstParDirPath
    If bCHOOSE_DOWNLOAD_DIR_PATH = True Then
        If bCHOOSE_FILE_AT_DIALOG_BOX = True Then
            '�t�H���_�I���_�C�A���O�\����BrowseForFolder(Shell.Application)
            '  �������p�X���w��ł��Ȃ����߁A�g�p���Ȃ�
            'Dim objFolder
            'Set objFolder = CreateObject("Shell.Application").BrowseForFolder(0, "�o�͐�t�H���_��I�����Ă�������", &H200, "c:\")
            'sDstParDirPath = objFolder
            
            '�t�H���_�I���_�C�A���O�\����FileDialog(Excel.Application)
            sDstParDirPath = ShowFolderSelectDialog( sIniDstParDirPath, "" )
        Else
            sDstParDirPath = InputBox( "�t�H���_��I�����Ă�������", sPROG_NAME, sIniDstParDirPath )
        End If
    Else
        sDstParDirPath = objWshShell.SpecialFolders("Desktop")
    End If
    
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sDstParDirPath)
    
    If objFSO.FolderExists( sDstParDirPath ) = False Then '�L�����Z���̏ꍇ
        MsgBox "���s���L�����Z������܂����B", vbOKOnly, sPROG_NAME
        Exit Sub
    End If
    
    '*** �ޔ�p�t�H���_�쐬 ***
    Dim sDstParEvaDirPath
    If bEVACUATE_ORG_FILE = True Then
        sDstParEvaDirPath = sDstParDirPath & "\_#" & sORIGINAL_FILE_PREFIX & "#"
        '�t�H���_�쐬
        If objFSO.FolderExists( sDstParEvaDirPath ) = False Then
            objFSO.CreateFolder( sDstParEvaDirPath )
        End If
    Else
        sDstParEvaDirPath = sDstParDirPath
    End If
    
    '*** �V���[�g�J�b�g�쐬 ***
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        '�t�@�C��/�t�H���_����
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        
        Dim sAddDate
        Dim sDstOrgFilePath
        Dim sDstCpyFilePath
        Dim sDstShrtctFilePath
        
        '### �t�@�C�� ###
        If bFolderExists = False And bFileExists = True Then
            '�ǉ�������擾�����`
            If lADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE )
            ElseIf lADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified, lDATE_STR_TYPE )
                Set objFile = Nothing
            Else
                MsgBox "�ulADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, sPROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sSrcFileName
            Dim sSrcFileExt
            sSrcFileName = objFSO.GetFileName( sSelectedPath )
            sSrcFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcFileName & ".#" & sEDIT_FILE_PREFIX     & "#" & sAddDate & "." & sSrcFileExt
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcFileName & ".#" & sORIGINAL_FILE_PREFIX & "#" & sAddDate & "." & sSrcFileExt
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcFileName & ".#" & sSHORTCUT_FILE_SUFFIX & "#" & sAddDate & ".lnk"
            
            '�t�@�C���R�s�[
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = objFSO.GetParentFolderName( sSelectedPath )
                .Save
            End With
            
        '### �t�H���_ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '�ǉ�������擾�����`
            If lADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE )
            ElseIf lADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified, lDATE_STR_TYPE )
                Set objFolder = Nothing
            Else
                MsgBox "�ulADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, sPROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sSrcDirName
            sSrcDirName = objFSO.GetFileName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcDirName & ".#" & sEDIT_FILE_PREFIX     & "#" & sAddDate
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcDirName & ".#" & sORIGINAL_FILE_PREFIX & "#" & sAddDate
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcDirName & ".#" & sSHORTCUT_FILE_SUFFIX & "#" & sAddDate & ".lnk"
            
            '�t�H���_�R�s�[
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = objFSO.GetParentFolderName( sSelectedPath )
                .Save
            End With
            
        '### �t�@�C��/�t�H���_�ȊO ###
        Else
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, sPROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
            Exit Sub
        End If
    Next
    
    '*** �t�H���_���J�� ***
    CreateObject("Shell.Application").Explore sDstParDirPath
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
End Sub

'===============================================================================
'= �e�X�g�֐�
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    MsgBox "=== test start ==="
    Select Case lTestCase
        Case 1
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
