'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )          ' ShowFolderSelectDialog()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )              ' ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\SettingFileClass.vbs" )    ' SettingFile

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sPROG_NAME = "参照ファイル複製"
Const lADD_DATE_TYPE = 1 '付与する日時の種別（1:現在日時、2:ファイル/フォルダ更新日時）
Const lDATE_STR_TYPE = 1
Const bEVACUATE_ORG_FILE = True
Const bCHOOSE_DOWNLOAD_DIR_PATH = False
Const bCHOOSE_FILE_AT_DIALOG_BOX = True
Const sSHORTCUT_FILE_SUFFIX = "s"
Const sORIGINAL_FILE_PREFIX = "o"
Const sEDIT_FILE_PREFIX     = "e"
Const sTEMP_FILE_NAME = "CopyRefFile.cfg"

'===============================================================================
'= 本処理
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
'= メイン関数
'===============================================================================
Public Sub Main()
    Dim sSrcParDirPath
    Dim sIniDstParDirPath
    Dim cSelectedPaths
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    
    '*** 選択ファイル取得 ***
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In cArgs
            cSelectedPaths.add sArg
            If sSrcParDirPath = "" Then
                sSrcParDirPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sSrcParDirPath = WScript.Env("Current")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sSrcParDirPath = "C:\Users\draem\OneDrive\デスクトップ"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\codes\vbs\tools\win\file_ope\CpyAndAddModDate.vbs"
        cSelectedPaths.Add "C:\codes\vbs\tools\win\file_ope\CpyAndAddNowDate.vbs"
        cSelectedPaths.Add "C:\codes\vbs\tools\win\other"
    End If
    
    '*** ファイルパスチェック ***
    If cSelectedPaths.Count = 0 Then
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    
    '*** 上書き確認 ***
    '実行速度を高めるため、上書き確認省略
    'Dim vbAnswer
    'vbAnswer = MsgBox( "既にファイルが存在している場合、上書きされます。実行しますか？", vbOkCancel, sPROG_NAME )
    'If vbAnswer = vbOk Then
    '    'Do Nothing
    'Else
    '    MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
    '    Exit Sub
    'End If
    
    '*** 出力先選択 ***
    '出力先フォルダパス取得 from クリップボード
    'sIniDstParDirPath = CreateObject("htmlfile").ParentWindow.Clipboarddata.GetData("text")
    'If objFSO.FolderExists( sIniDstParDirPath ) = False Then
    '    sIniDstParDirPath = objWshShell.SpecialFolders("Desktop")
    'End If
    
    '出力先フォルダパス取得 from 設定ファイル
    Dim clSetting
    Set clSetting = New SettingFile
    Dim sSettingFilePath
    sSettingFilePath = objFSO.GetSpecialFolder(2) & "\" & sTEMP_FILE_NAME
    
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
            sDstParDirPath = ShowFolderSelectDialog( sIniDstParDirPath, "" )
        Else
            sDstParDirPath = InputBox( "フォルダを選択してください", sPROG_NAME, sIniDstParDirPath )
        End If
    Else
        sDstParDirPath = objWshShell.SpecialFolders("Desktop")
    End If
    
    Call clSetting.WriteItemToFile(sSettingFilePath, "sDST_PAR_DIR_PATH", sDstParDirPath)
    
    If objFSO.FolderExists( sDstParDirPath ) = False Then 'キャンセルの場合
        MsgBox "実行がキャンセルされました。", vbOKOnly, sPROG_NAME
        Exit Sub
    End If
    
    '*** 退避用フォルダ作成 ***
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
    
    '*** ショートカット作成 ***
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        'ファイル/フォルダ判定
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        
        Dim sAddDate
        Dim sDstOrgFilePath
        Dim sDstCpyFilePath
        Dim sDstShrtctFilePath
        
        '### ファイル ###
        If bFolderExists = False And bFileExists = True Then
            '追加文字列取得＆整形
            If lADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE )
            ElseIf lADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified, lDATE_STR_TYPE )
                Set objFile = Nothing
            Else
                MsgBox "「lADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, sPROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sSrcFileName
            Dim sSrcFileExt
            sSrcFileName = objFSO.GetFileName( sSelectedPath )
            sSrcFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcFileName & ".#" & sEDIT_FILE_PREFIX     & "#" & sAddDate & "." & sSrcFileExt
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcFileName & ".#" & sORIGINAL_FILE_PREFIX & "#" & sAddDate & "." & sSrcFileExt
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcFileName & ".#" & sSHORTCUT_FILE_SUFFIX & "#" & sAddDate & ".lnk"
            
            'ファイルコピー
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = objFSO.GetParentFolderName( sSelectedPath )
                .Save
            End With
            
        '### フォルダ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '追加文字列取得＆整形
            If lADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), lDATE_STR_TYPE )
            ElseIf lADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified, lDATE_STR_TYPE )
                Set objFolder = Nothing
            Else
                MsgBox "「lADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, sPROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sSrcDirName
            sSrcDirName = objFSO.GetFileName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcDirName & ".#" & sEDIT_FILE_PREFIX     & "#" & sAddDate
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcDirName & ".#" & sORIGINAL_FILE_PREFIX & "#" & sAddDate
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcDirName & ".#" & sSHORTCUT_FILE_SUFFIX & "#" & sAddDate & ".lnk"
            
            'フォルダコピー
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = objFSO.GetParentFolderName( sSelectedPath )
                .Save
            End With
            
        '### ファイル/フォルダ以外 ###
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, sPROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
            Exit Sub
        End If
    Next
    
    '*** フォルダを開く ***
    CreateObject("Shell.Application").Explore sDstParDirPath
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
End Sub

'===============================================================================
'= テスト関数
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
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}
