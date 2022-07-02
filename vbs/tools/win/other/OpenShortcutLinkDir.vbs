'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "ショートカットファイルの指示先フォルダに移動"

Dim bIsContinue
bIsContinue = True

Dim sFilePath

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
       sFilePath = WScript.Arguments(0)
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sFilePath = WScript.Env("Focused")
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sFilePath = "C:\codes\vbs\tools\win\other\test.vbs - ショートカット.lnk"
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If bIsContinue = True Then
    If sFilePath = "" Then
        MsgBox "ファイルが選択されていません", vbYes, sPROG_NAME
        MsgBox "処理を中断します", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
    If objFSO.GetExtensionName( sFilePath ) <> "lnk" Then
        MsgBox "ショートカットファイルを選択してください" & vbNewLine & sFilePath, vbYes, sPROG_NAME
        MsgBox "処理を中断します", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ショートカット指示先フォルダオープン ***
Dim objShell
Set objShell = WScript.CreateObject("Shell.Application")
If bIsContinue = True Then
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim sTrgtFilePath
    Dim sTrgtDirPath
    sTrgtFilePath = objWshShell.CreateShortcut( sFilePath ).TargetPath
    sTrgtDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        objShell.Explore sTrgtDirPath
        Set objShell = Nothing
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        WScript.Open(sFilePath)
    Else 'デバッグ実行
        MsgBox "「" & sTrgtDirPath & "」をエクスプローラーで開きます"
        objShell.Explore sTrgtDirPath
        Set objShell = Nothing
    End If
Else
    'Do Nothing
End If
