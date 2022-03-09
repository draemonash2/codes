Option Explicit

'####################################################################
'### 設定
'####################################################################
Const lDATE_STR_TYPE = 3
Const bEVACUATE_ORG_FILE = True
Const bCHOOSE_DOWNLOAD_DIR_PATH = False
Const bCHOOSE_FILE_AT_DIALOG_BOX = True
Const sSHORTCUT_FILE_SUFFIX = "s"
Const sORIGINAL_FILE_PREFIX = "o"
Const sEDIT_FILE_PREFIX     = "e"
Const sTEMP_FILE_NAME = "CopyAsWorkFileFromURL.cfg"

'####################################################################
'### インクルード
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )         ' ShowFolderSelectDialog()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )             ' ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\SettingFileClass.vbs" )   ' SettingFile
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )                ' DownloadFile()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "ファイルダウンロード"

Dim bIsContinue
bIsContinue = True

Dim sSrcParDirPath
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim clSetting
Dim sSettingFilePath
sSettingFilePath = objFSO.GetSpecialFolder(2) & "\" & sTEMP_FILE_NAME

'*** ダウンロードファイルURL入力 ***
'出力先フォルダパス取得 from 設定ファイル
If bIsContinue = True Then
    Dim sIniDownloadFileUrl
    Set clSetting = New SettingFile
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDOWNLOAD_FILE_PATH", sIniDownloadFileUrl, "", False)
    Dim sDownloadFileUrl
    sDownloadFileUrl = InputBox( "ダウンロードしたいファイルのURLを入力してください", sPROG_NAME, sIniDownloadFileUrl )
    If IsEmpty(sDownloadFileUrl) = True Then
        MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sPROG_NAME
        bIsContinue = False
    ElseIf sDownloadFileUrl = "" Then
        MsgBox "URLが入力されていないため、処理を中断します。", vbExclamation, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDOWNLOAD_FILE_PATH", sDownloadFileUrl)
    Set clSetting = Nothing
End If

'*** 取得元URL入力 ***
'出力先フォルダパス取得 from 設定ファイル
If bIsContinue = True Then
    Dim sIniDownloadSrcUrl
    Set clSetting = New SettingFile
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDOWNLOAD_SRC_URL", sIniDownloadSrcUrl, "", False)
    Dim sDownloadSrcUrl
    sDownloadSrcUrl = InputBox( "ダウンロード元のURLを入力してください", sPROG_NAME, sIniDownloadSrcUrl )
    If IsEmpty(sDownloadSrcUrl) = True Then
        MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sPROG_NAME
        bIsContinue = False
    ElseIf sDownloadSrcUrl = "" Then
        MsgBox "URLが入力されていないため、処理を中断します。", vbExclamation, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDOWNLOAD_SRC_URL", sDownloadSrcUrl)
    Set clSetting = Nothing
End If

'*** 出力先選択 ***
'出力先フォルダパス取得 from 設定ファイル
If bIsContinue = True Then
    Dim sIniDstParDirPath
    Set clSetting = New SettingFile
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sIniDstParDirPath, objWshShell.SpecialFolders("Desktop"), False)
    Dim sDstParDirPath
    If bCHOOSE_DOWNLOAD_DIR_PATH = True Then
        If bCHOOSE_FILE_AT_DIALOG_BOX = True Then
            'フォルダ選択ダイアログ表示＠BrowseForFolder(Shell.Application)
            
            '  →初期パスを指定できないため、使用しない
            'Dim objFolder
            'Set objFolder = CreateObject("Shell.Application").BrowseForFolder(0, "出力先フォルダを選択してください", &H200, "c:\")
            'sDstParDirPath = objFolder
            
            'フォルダ選択ダイアログ表示＠FileDialog(Excel.Application)
            sDstParDirPath = ShowFolderSelectDialog( sIniDstParDirPath, sPROG_NAME )
        Else
            sDstParDirPath = InputBox( "フォルダを選択してください", sPROG_NAME, sIniDstParDirPath )
        End If
    Else
        sDstParDirPath = objWshShell.SpecialFolders("Desktop")
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sDstParDirPath)
    Set clSetting = Nothing
End If

If bIsContinue = True Then
    If objFSO.FolderExists( sDstParDirPath ) = False Then 'キャンセルの場合
        MsgBox "実行がキャンセルされました。", vbExclamation, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
End If

'*** 退避用フォルダ作成 ***
If bIsContinue = True Then
    Dim sDstParEvaDirPath
    If bEVACUATE_ORG_FILE = True Then
        sDstParEvaDirPath = sDstParDirPath & "\_#" & sORIGINAL_FILE_PREFIX & "#"
        'フォルダ作成
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
    sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE ) '追加文字列取得＆整形
    Dim sDownloadFileName
    Dim sDownloadFileBaseName
    Dim sDownloadFileExt
    sDownloadFileName = objFSO.GetFileName( sDownloadFileUrl )
    sDownloadFileBaseName = objFSO.GetBaseName( sDownloadFileName )
    sDownloadFileExt = objFSO.GetExtensionName( sDownloadFileName )
    
    sDstCpyFilePath    = sDstParDirPath    & "\" & sDownloadFileName & "_#" & sEDIT_FILE_PREFIX     & sAddDate & "#." & sDownloadFileExt
    sDstOrgFilePath    = sDstParEvaDirPath & "\" & sDownloadFileName & "_#" & sORIGINAL_FILE_PREFIX & sAddDate & "#." & sDownloadFileExt
    sDstShrtctFilePath = sDstParEvaDirPath & "\" & sDownloadFileName & "_#" & sSHORTCUT_FILE_SUFFIX & sAddDate & "#.lnk"
    
    '*** ファイルダウンロード ***
    Call DownloadFile(sDownloadFileUrl, sDstCpyFilePath )
    
    '*** ファイルコピー ***
    objFSO.CopyFile sDstCpyFilePath, sDstOrgFilePath, True
    
    '*** ショートカット作成 ***
    With objWshShell.CreateShortcut( sDstShrtctFilePath )
        .TargetPath = sDownloadSrcUrl
        .Save
    End With
    
    '*** フォルダを開く ***
    CreateObject("Shell.Application").Explore sDstParDirPath
End If

Set objFSO = Nothing
Set clSetting = Nothing
Set objWshShell = Nothing

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

