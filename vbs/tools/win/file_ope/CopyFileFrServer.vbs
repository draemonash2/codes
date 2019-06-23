'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const ADD_DATE_TYPE = 1 '付与する日時の種別（1:現在日時、2:ファイル/フォルダ更新日時）
Const EVACUATE_ORG_FILE = True
Const SHORTCUT_FILE_SUFFIX = "#s#"
Const ORIGINAL_FILE_PREFIX = "#o#"
Const EDIT_FILE_PREFIX     = "#e#"

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ショートカット＆コピーファイル作成"

Dim bIsContinue
bIsContinue = True

Dim sOrgDirPath
Dim cSelectedPaths
Dim objFSO

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
            If sOrgDirPath = "" Then
                sOrgDirPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sOrgDirPath = WScript.Env("Current")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sOrgDirPath = "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H20年度 ゼミ出席簿.xls"
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H21年度 ゼミ出席簿.xls"
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** 上書き確認 ***
'実行速度を高めるため、上書き確認省略
'If bIsContinue = True Then
'    Dim vbAnswer
'    vbAnswer = MsgBox( "既にファイルが存在している場合、上書きされます。実行しますか？", vbOkCancel, PROG_NAME )
'    If vbAnswer = vbOk Then
'        'Do Nothing
'    Else
'        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
'        bIsContinue = False
'    End If
'Else
'    'Do Nothing
'End If

'*** 出力先選択 ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = ShowFolderSelectDialog( sOrgDirPath )

    If objFSO.FolderExists( sDstParDirPath ) = False Then 'キャンセルの場合
        MsgBox "実行がキャンセルされました。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** 退避用フォルダ作成 ***
If bIsContinue = True Then
    Dim sDstParEvaDirPath
    If EVACUATE_ORG_FILE = True Then
        sDstParEvaDirPath = sDstParDirPath & "\_" & ORIGINAL_FILE_PREFIX
        'フォルダ作成
        If objFSO.FolderExists( sDstParEvaDirPath ) = False Then
            objFSO.CreateFolder( sDstParEvaDirPath )
        End If
    Else
        sDstParEvaDirPath = sDstParDirPath
    End If
Else
    'Do Nothing
End If

'*** ショートカット作成 ***
If bIsContinue = True Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
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
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified )
                Set objFile = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sSrcFileName
            Dim sSrcFileBaseName
            Dim sSrcFileExt
            sSrcFileName = objFSO.GetFileName( sSelectedPath )
            sSrcFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sSrcFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcFileName & "_" & EDIT_FILE_PREFIX     & sAddDate & "." & sSrcFileExt
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcFileName & "_" & ORIGINAL_FILE_PREFIX & sAddDate & "." & sSrcFileExt
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcFileName & "_" & SHORTCUT_FILE_SUFFIX & sAddDate & ".lnk"
            
            'ファイルコピー
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### フォルダ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '追加文字列取得＆整形
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified )
                Set objFolder = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sSrcDirName
            sSrcDirName = objFSO.GetFileName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcDirName & "_" & EDIT_FILE_PREFIX     & sAddDate
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcDirName & "_" & ORIGINAL_FILE_PREFIX & sAddDate
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcDirName & "_" & SHORTCUT_FILE_SUFFIX & sAddDate & ".lnk"
            
            'フォルダコピー
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### ファイル/フォルダ以外 ###
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "ショートカット＆コピーファイルの作成が完了しました！", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

' ==================================================================
' = 概要    日時文字列をファイル/フォルダ名に適用できる形式に変換する
' = 引数    sDateRaw    String  [in]    日時（例：2017/8/5 12:59:58）
' = 戻値                String          日時（例：20170805-125958）
' = 覚書    なし
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateRaw _
)
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
    sTargetStr = sDateRaw
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '検索パターンを設定
    oRegExp.IgnoreCase = True                       '大文字と小文字を区別しない
    oRegExp.Global = True                           '文字列全体を検索
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  'パターンマッチ実行
    Dim sDateStr
    With oMatchResult(0)
        sDateStr = Right( .SubMatches(0), 2 ) & _
                   String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
                   String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
                   String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
                   String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4)
    End With
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    ConvDate2String = sDateStr
End Function

' ==================================================================
' = 概要    フォルダ選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトフォルダパス
' = 戻値                String        フォルダ選択結果
' = 覚書    ・存在しないフォルダパスを選択した場合、空文字列を返却する
' =         ・キャンセルを押下した場合、空文字列を返却する
' ==================================================================
Private Function ShowFolderSelectDialog( _
    ByVal sInitPath _
)
    Const msoFileDialogFolderPicker = 4
    Const xlMinimized = -4140
    
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False '非表示にしても閉じる際にちらっと表示されちゃう。
    objExcel.WindowState = xlMinimized '上記理由から最小化もしとく。
    
    Dim fdDialog
    Set fdDialog = objExcel.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    'ダイアログ表示
    Dim lResult
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFolderSelectDialog = ""
    Else
        Dim sSelectedPath
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists( sSelectedPath ) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
'   Call Test_ShowFolderSelectDialog()
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        
        Dim sInitPath
        sInitPath = objWshShell.SpecialFolders("Desktop")
        'sInitPath = ""
        
        MsgBox ShowFolderSelectDialog( sInitPath )
    End Sub

Public Function SetFileAttributes( _
    ByVal sFilePath, _
    ByVal sDateRaw _
)
    Const SET_ATTR_READONLY = 1     ' 読み取り専用ファイル
    Const SET_ATTR_HIDDEN   = 2     ' 隠しファイル
    Const SET_ATTR_SYSTEM   = 4     ' システム・ファイル
    Const SET_ATTR_ARCHIVE  = 32    ' 前回のバックアップ以降に変更されていれば1
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile("test.txt ")
End Function

