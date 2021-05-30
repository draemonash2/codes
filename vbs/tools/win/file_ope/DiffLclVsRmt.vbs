Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" ) 'DownloadFile()

'===============================================================================
'= 本処理
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 4 Then
    MsgBox "引数を正しく指定してください。", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sUserName
Dim sPassword
Dim sLoginServerName
Dim sHomeDirPath
sUserName = WScript.Arguments(0)
sPassword = WScript.Arguments(1)
sLoginServerName = WScript.Arguments(2)
sHomeDirPath = WScript.Arguments(3)

Dim sOutputMsg
sOutputMsg = WScript.ScriptName

'=== 事前処理 ===
Dim sDownloadTrgtDirPath
Dim sScpProgramPath
Dim sDiffProgramPath
Dim sCodesDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sScpProgramPath = GetEnvVariable("MYEXEPATH_WINSCP")
'sScpProgramPath = "C:\Users\draem\Programs\program\prg_exe\WinSCP\WinSCP.exe" ★debug
sDiffProgramPath = GetEnvVariable("MYEXEPATH_WINMERGE")
sCodesDirPath = GetEnvVariable("MYDIRPATH_CODES")

'=== ファイル受信(Remote → Local) ===
Dim vAnswer
vAnswer = MsgBox("ダウンロードを開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbOk Then
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.vimrc "     & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.bashrc "    & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.inputrc "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
Else
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== ファイルバックアップ ===
objFSO.CopyFile sDownloadTrgtDirPath & "\.vimrc",   sDownloadTrgtDirPath & "\.vimrc_rmtorg"
objFSO.CopyFile sDownloadTrgtDirPath & "\.bashrc",  sDownloadTrgtDirPath & "\.bashrc_rmtorg"
objFSO.CopyFile sDownloadTrgtDirPath & "\.inputrc", sDownloadTrgtDirPath & "\.inputrc_rmtorg"

'=== フォルダ比較 ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.inputrc"    & """ """ & sDownloadTrgtDirPath & "\.inputrc"  & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.bashrc"     & """ """ & sDownloadTrgtDirPath & "\.bashrc"   & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\.vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\_vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\_gvimrc"       & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
vAnswer = MsgBox("比較/マージが完了したらOKを押してください。", vbOkOnly, sOutputMsg)

'=== ファイル送信(Local → Remote) ===
vAnswer = MsgBox("編集したファイルを送信しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.vimrc" & """ ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.bashrc" & """ ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.inputrc" & """ ""exit""", 0, True
End If

'=== ファイル削除 ===
vAnswer = MsgBox("ダウンロードしたファイルを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.vimrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.vimrc_rmtorg"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.bashrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.bashrc_rmtorg"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.inputrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.inputrc_rmtorg"
End If

MsgBox "処理が完了しました！", vbYesNo, sOutputMsg

'===============================================================================
'= 内部関数
'===============================================================================
' ==================================================================
' = 概要    環境変数を取得する
' = 引数    sEnvVar     String  [in]    環境変数名
' = 戻値                String          環境変数値
' = 覚書    ・環境変数が存在しない場合、処理を中断する
' = 依存    なし
' = 所属    本スクリプト
' ==================================================================
Private Function GetEnvVariable( _
    ByVal sEnvVar _
)
    Dim sGetValue
    sGetValue = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%" & sEnvVar & "%")
    If InStr(sGetValue, "%") > 0 then
        MsgBox "環境変数「" & sEnvVar & "」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
        WScript.Quit
    End If
    GetEnvVariable = sGetValue
End Function

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

