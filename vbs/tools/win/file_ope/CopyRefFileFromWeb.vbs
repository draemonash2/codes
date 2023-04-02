Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )         ' ShowFolderSelectDialog()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )             ' ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\SettingFileClass.vbs" )   ' SettingFile
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )                ' DownloadFile()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const bEXEC_TEST = False '�e�X�g�p
Const sPROG_NAME = "Web����Q�ƃt�@�C������"
Const lDATE_STR_TYPE = 1
Const bEVACUATE_ORG_FILE = True
Const bCHOOSE_DOWNLOAD_DIR_PATH = False
Const bCHOOSE_FILE_AT_DIALOG_BOX = True
Const sSHORTCUT_FILE_SUFFIX = "s"
Const sORIGINAL_FILE_PREFIX = "o"
Const sEDIT_FILE_PREFIX     = "e"
Const sTEMP_FILE_NAME = "CopyRefFileFromWeb.cfg"

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
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim clSetting
    Dim sSettingFilePath
    sSettingFilePath = objFSO.GetSpecialFolder(2) & "\" & sTEMP_FILE_NAME
    
    '*** �_�E�����[�h�t�@�C��URL���� ***
    '�o�͐�t�H���_�p�X�擾 from �ݒ�t�@�C��
    Dim sIniDownloadFileUrl
    Set clSetting = New SettingFile
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDOWNLOAD_FILE_PATH", sIniDownloadFileUrl, "", False)
    Dim sDownloadFileUrl
    sDownloadFileUrl = InputBox( "�_�E�����[�h�������t�@�C����URL����͂��Ă�������", sPROG_NAME, sIniDownloadFileUrl )
    If IsEmpty(sDownloadFileUrl) = True Then
        MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sPROG_NAME
        Exit Sub
    ElseIf sDownloadFileUrl = "" Then
        MsgBox "URL�����͂���Ă��Ȃ����߁A�����𒆒f���܂��B", vbExclamation, sPROG_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDOWNLOAD_FILE_PATH", sDownloadFileUrl)
    Set clSetting = Nothing
    
    '*** �擾��URL���� ***
    '�o�͐�t�H���_�p�X�擾 from �ݒ�t�@�C��
    Dim sIniDownloadSrcUrl
    Set clSetting = New SettingFile
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDOWNLOAD_SRC_URL", sIniDownloadSrcUrl, "", False)
    Dim sDownloadSrcUrl
    sDownloadSrcUrl = InputBox( "�_�E�����[�h����URL����͂��Ă�������", sPROG_NAME, sIniDownloadSrcUrl )
    If IsEmpty(sDownloadSrcUrl) = True Then
        MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sPROG_NAME
        Exit Sub
    ElseIf sDownloadSrcUrl = "" Then
        MsgBox "URL�����͂���Ă��Ȃ����߁A�����𒆒f���܂��B", vbExclamation, sPROG_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDOWNLOAD_SRC_URL", sDownloadSrcUrl)
    Set clSetting = Nothing
    
    '*** �o�͐�I�� ***
    '�o�͐�t�H���_�p�X�擾 from �ݒ�t�@�C��
    Dim sIniDstParDirPath
    Set clSetting = New SettingFile
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
            sDstParDirPath = ShowFolderSelectDialog( sIniDstParDirPath, sPROG_NAME )
        Else
            sDstParDirPath = InputBox( "�t�H���_��I�����Ă�������", sPROG_NAME, sIniDstParDirPath )
        End If
    Else
        sDstParDirPath = objWshShell.SpecialFolders("Desktop")
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sDstParDirPath)
    Set clSetting = Nothing
    
    If objFSO.FolderExists( sDstParDirPath ) = False Then '�L�����Z���̏ꍇ
        MsgBox "���s���L�����Z������܂����B", vbExclamation, sPROG_NAME
        Exit Sub
    Else
        'Do Nothing
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
    
    Dim sAddDate
    Dim sDstOrgFilePath
    Dim sDstCpyFilePath
    Dim sDstShrtctFilePath
    sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE ) '�ǉ�������擾�����`
    Dim sDownloadFileName
    Dim sDownloadFileBaseName
    Dim sDownloadFileExt
    sDownloadFileName = objFSO.GetFileName( sDownloadFileUrl )
    sDownloadFileBaseName = objFSO.GetBaseName( sDownloadFileName )
    sDownloadFileExt = objFSO.GetExtensionName( sDownloadFileName )
    
    sDstCpyFilePath    = sDstParDirPath    & "\" & sDownloadFileName & ".#" & sEDIT_FILE_PREFIX     & "#" & sAddDate & "." & sDownloadFileExt
    sDstOrgFilePath    = sDstParEvaDirPath & "\" & sDownloadFileName & ".#" & sORIGINAL_FILE_PREFIX & "#" & sAddDate & "." & sDownloadFileExt
    sDstShrtctFilePath = sDstParEvaDirPath & "\" & sDownloadFileName & ".#" & sSHORTCUT_FILE_SUFFIX & "#" & sAddDate & ".lnk"
    
    '*** �t�@�C���_�E�����[�h ***
    Call DownloadFile(sDownloadFileUrl, sDstCpyFilePath )
    
    '*** �t�@�C���R�s�[ ***
    objFSO.CopyFile sDstCpyFilePath, sDstOrgFilePath, True
    
    '*** �V���[�g�J�b�g�쐬 ***
    With objWshShell.CreateShortcut( sDstShrtctFilePath )
        .TargetPath = sDownloadSrcUrl
        .Save
    End With
    
    '*** �t�H���_���J�� ***
    CreateObject("Shell.Application").Explore sDstParDirPath
    
    Set objFSO = Nothing
    Set clSetting = Nothing
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