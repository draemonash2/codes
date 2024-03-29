'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "カレントフォルダ配下の特定ファイルを Vim で全て開く"

Dim bIsContinue
bIsContinue = True

Dim sCurDirPath

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        sCurDirPath = InputBox( "フォルダパスを指定してください", sPROG_NAME, WScript.Arguments(0) )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sCurDirPath = WScript.Env("Current")
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sCurDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** フォルダ存在確認 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists( sCurDirPath ) Then
        ' Do Nothing
    Else
        MsgBox "指定されたフォルダが存在しません。" & vbNewLine & sCurDirPath, vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        WScript.Quit
    End If
End If

'*** 拡張子選択 ***
If bIsContinue = True Then
    Dim sExtNames
    sExtNames = InputBox( _
        "拡張子を選択してください。" & vbNewLine & _
        "複数の拡張子を指定する時はスペースで区切ります。" & vbNewLine & _
        "  例１）*.txt *.c" & vbNewLine & _
        "  例２）*.*" & vbNewLine & _
        "" , _
        "title", _
        "*.c *.h" _
    )
    If sExtNames = "" Then
        MsgBox "拡張子が選択されていません", vbYes, sPROG_NAME
        MsgBox "処理を中断します", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ファイルリスト作成 ***
If bIsContinue = True Then
    'ファイルリスト出力コマンド実行
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = objWshShell.SpecialFolders("Templates") & "\open_file_list.txt"
    'MsgBox sTmpFilePath '★DEBUG★
    sExecCmd = "cd """ & sCurDirPath & """ & dir " & sExtNames & " /b /s /a:a-d > """ & sTmpFilePath & """"
    'MsgBox sExecCmd '★DEBUG★
    objWshShell.Run "cmd /c" & sExecCmd, 0, True
    
    '出力したファイルリスト取り込み
    Dim objFile
    Dim sTextAll
    On Error Resume Next
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    Dim asFileList
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
        'MsgBox Err.Number '★DEBUG★
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
            asFileList = Split( sTextAll, vbNewLine )
            objFile.Close
        Else
            MsgBox "エラーが発生しました。 [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, sPROG_NAME
            MsgBox "処理を中断します", vbYes, sPROG_NAME
            bIsContinue = False
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        MsgBox "エラーが発生しました。 [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, sPROG_NAME
        MsgBox "処理を中断します。", vbYes, sPROG_NAME
        bIsContinue = False
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error Goto 0
    'MsgBox Ubound(asFileList) '★DEBUG★
Else
    'Do Nothing
End If

'*** ファイルオープン実行 ***
If bIsContinue = True Then
    Dim sFilePathList
    sFilePathList = """"
    Dim lIdx
    lIdx = 0
    For Each sFilePath In asFileList
        If lIdx = 0 Then
            sFilePathList = """" & sFilePath & """"
        Else
            sFilePathList = sFilePathList & " """ & sFilePath & """"
        End If
        lIdx = lIdx + 1
    Next
    'MsgBox sFilePathList '★DEBUG★
    
    Dim sExePath
    sExePath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_GVIM%")
    If InStr(sExePath, "%") > 0 then
        MsgBox "環境変数が設定されていません。" & vbNewLine & "処理を中断します。", vbYes, sPROG_NAME
        WScript.Quit
    End If
    
    objWshShell.Run "cmd /c " & sExePath & " " & sFilePathList, 0, False
Else
    'Do Nothing
End If
