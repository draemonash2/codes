Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )        'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'CreateDirectry()

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
Dim sDownloadTrgtDirPathRaw
Dim sDownloadTrgtDirPath
Dim sScpProgramPath
Dim sDiffProgramPath
Dim sCodesDirPath
sDownloadTrgtDirPathRaw = objWshShell.SpecialFolders("Desktop")
sScpProgramPath = GetEnvVariable("MYEXEPATH_WINSCP")
sDiffProgramPath = GetEnvVariable("MYEXEPATH_WINMERGE")
sCodesDirPath = GetEnvVariable("MYDIRPATH_CODES")

'バックアップフォルダ作成
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = sDownloadTrgtDirPathRaw & "\" & sDateSuffix & "\" & sLoginServerName
Call CreateDirectry( sDownloadTrgtDirPath )

Dim vAnswer
vAnswer = MsgBox("ファイルを受信します。ダウンロードを開始しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== ファイル受信(Remote → Local) ===
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.vimrc "     & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.bashrc "    & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.inputrc "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.gdbinit "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    
    '=== ファイルバックアップ ===
    objFSO.CopyFile sDownloadTrgtDirPath & "\.vimrc",   sDownloadTrgtDirPath & "\.vimrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.bashrc",  sDownloadTrgtDirPath & "\.bashrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.inputrc", sDownloadTrgtDirPath & "\.inputrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.gdbinit", sDownloadTrgtDirPath & "\.gdbinit_rmtorg"
    
    '=== フォルダ比較 ===
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.gdbinit"    & """ """ & sDownloadTrgtDirPath & "\.gdbinit"  & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.inputrc"    & """ """ & sDownloadTrgtDirPath & "\.inputrc"  & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.bashrc"     & """ """ & sDownloadTrgtDirPath & "\.bashrc"   & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\.vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
    vAnswer = MsgBox("比較/マージが完了したらOKを押してください。", vbOkOnly, sOutputMsg)
Else
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

vAnswer = MsgBox("編集したファイルを送信しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== ファイル送信(Local → Remote) ===
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.vimrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.bashrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.inputrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.gdbinit" & """ ""exit""", 0, True
End If

vAnswer = MsgBox("ダウンロードしたファイルを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== フォルダ削除 ===
    objFSO.DeleteFolder sDownloadTrgtDirPathRaw & "\" & sDateSuffix, True
End If

'MsgBox "処理が完了しました！", vbYesNo, sOutputMsg

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

