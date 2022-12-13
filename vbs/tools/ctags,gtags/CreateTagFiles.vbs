'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "タグファイル作成"

Dim bIsContinue
bIsContinue = True

Dim sTrgtDirPath

If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        sTrgtDirPath = InputBox( "フォルダパスを指定してください", sPROG_NAME, WScript.Arguments(0) )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sTrgtDirPath = WScript.Env("Current")
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sTrgtDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** フォルダ存在確認 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists( sTrgtDirPath ) Then
        ' Do Nothing
    Else
        MsgBox "指定されたフォルダが存在しません。" & vbNewLine & sTrgtDirPath, vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        WScript.Quit
    End If
End If

'*** タグファイル作成 ***
If bIsContinue = True Then
    MsgBox "「ctags」と「gtags」にパスが通っていない場合は、パスを通してから実行してください。", vbOk, sPROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "タグファイルの作成が完了しました。", vbOk, sPROG_NAME
Else
    'Do Nothing
End If
