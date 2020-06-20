'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const ADD_DATE_TYPE = 1 '付与する日時の種別（1:現在日時、2:ファイル/フォルダ更新日時）
Const SHORTCUT_FILE_SUFFIX = "#Src#"
Const ORIGINAL_FILE_PREFIX = "#Org#"
Const COPY_FILE_PREFIX     = "#Cpy#"

'####################################################################
'### 事前処理
'####################################################################
Call Include( "C:\codes\vbs\_lib\String.vbs" ) 'ConvDate2String()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "ショートカット＆コピーファイル作成(SVN)"

Dim bIsContinue
bIsContinue = True

Dim sOrgDirPath
Dim cSelectedPaths
Dim objFSO

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Dim sDefaultPath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sOrgDirPath = InputBox( "ファイルパスを指定してください", sPROG_NAME, sDefaultPath )
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
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** 上書き確認 ***
If bIsContinue = True Then
    Dim vbAnswer
    vbAnswer = MsgBox( "既にファイルが存在している場合、上書きされます。実行しますか？", vbOkCancel, sPROG_NAME )
    If vbAnswer = vbOk Then
        'Do Nothing
    Else
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'*** 出力先選択 ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = InputBox( "出力先を入力してください。", sPROG_NAME, sOrgDirPath )
    If sDstParDirPath = "" Then 'キャンセルの場合
        MsgBox "実行がキャンセルされました。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
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
                sAddDate = ConvDate2String( Now(), 1 )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified, 1 )
                Set objFile = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, sPROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sOrgFileName
            Dim sOrgFileBaseName
            Dim sOrgFileExt
            sOrgFileName = objFSO.GetFileName( sSelectedPath )
            sOrgFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sOrgFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate & "." & sOrgFileExt
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate & "." & sOrgFileExt
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgFileName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            'ファイルコピー
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### フォルダ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '追加文字列取得＆整形
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), 1 )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified, 1 )
                Set objFolder = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, sPROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sOrgDirName
            sOrgDirName = objFSO.GetFileName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgDirName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            'フォルダコピー
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### ファイル/フォルダ以外 ###
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, sPROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "ショートカット＆コピーファイルの作成が完了しました！", vbOKOnly, sPROG_NAME
Else
    'Do Nothing
End If

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

