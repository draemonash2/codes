Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )        'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'CreateDirectry()
                                                            'MoveToTrushBox()

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

Dim cTrgtFileNames
Dim cTrgtDirNames
Set cTrgtFileNames = CreateObject("System.Collections.ArrayList")
Set cTrgtDirNames = CreateObject("System.Collections.ArrayList")
cTrgtDirNames.Add "vim"   : cTrgtFileNames.Add ".vimrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".bashrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".inputrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".gdbinit"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".tmux.conf"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".tigrc"

'バックアップフォルダ作成
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix & "\" & sLoginServerName
Call CreateDirectry( sDownloadTrgtDirPath )

Dim vAnswer
Dim iIdx
vAnswer = MsgBox("ファイルを受信します。ダウンロードを開始しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    On Error Resume Next 'リモートにファイルが存在しなくても無視する
    '=== ファイル受信(Remote → Local) ===
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/" & cTrgtFileNames(iIdx) & " " & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
        If objFSO.FileExists(sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx)) Then
            'Do Nothing
        Else
            objFSO.CopyFile sCodesDirPath & "\" & cTrgtDirNames(iIdx) & "\" & cTrgtFileNames(iIdx), sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx)
        End If
    Next
    '=== ファイルバックアップ ===
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objFSO.CopyFile sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx), sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & "_rmtorg"
    Next
    '=== フォルダ比較 ===
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objWshShell.Run """" & sDiffProgramPath & """ -r -s """ & sCodesDirPath & "\" & cTrgtDirNames(iIdx) & "\" & cTrgtFileNames(iIdx) & """ """ & sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & """", 10, False
    Next
    On Error Goto 0
    MsgBox "比較/マージが完了したらOKを押してください。", vbOkOnly, sOutputMsg
Else
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

vAnswer = MsgBox("編集したファイルを送信しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== ファイル送信(Local → Remote) ===
    On Error Resume Next 'ローカルにファイルが存在しなくても無視する
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & """ ""exit""", 0, True
    Next
    On Error Goto 0
End If

vAnswer = MsgBox("ダウンロードしたファイルを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== フォルダ削除 ===
    'objFSO.DeleteFolder sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix, True
    Call MoveToTrushBox(sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix)
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

