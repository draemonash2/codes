Option Explicit

'  [使い方]
'    SyncCodesToRemote.vbs <username> <key> <server> <home_dir_path> [<auth_type>]
'      <username>：ユーザ名
'      <key>：【パスワード認証時】パスワード 【公開鍵認証時】秘密鍵格納先(*.ppk)
'      <server>：接続先（e.g. 123.345.678.901:22）
'      <auth_type>：キー種別 (0:パスワード認証 1:公開鍵認証)。省略可。デフォルト=パスワード認証。

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )        'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'CreateDirectry()
                                                            'MoveToTrushBox()

'===============================================================================
'= 設定
'===============================================================================
Dim csSYNC_TRGT_FILES
Set csSYNC_TRGT_FILES = CreateObject("System.Collections.ArrayList")

csSYNC_TRGT_FILES.Add "_dotfiles\.vimrc"
csSYNC_TRGT_FILES.Add "_dotfiles\.bashrc"
csSYNC_TRGT_FILES.Add "_dotfiles\.inputrc"
csSYNC_TRGT_FILES.Add "_dotfiles\.gdbinit"
csSYNC_TRGT_FILES.Add "_dotfiles\.tmux.conf"
csSYNC_TRGT_FILES.Add "_dotfiles\.tigrc"
csSYNC_TRGT_FILES.Add "_dotfiles\.ai_agents\AGENTS.md"
csSYNC_TRGT_FILES.Add "_dotfiles\.ai_agents\skills\func-header-comment\SKILL.md"
csSYNC_TRGT_FILES.Add "_dotfiles\.ai_agents\skills\func-tree\SKILL.md"

'===============================================================================
'= 本処理
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sLinuxUserName
Dim sKey
Dim sLoginServerName
Dim sAuthType
If WScript.Arguments.Count >= 3 Then
    sLinuxUserName = WScript.Arguments(0)
    sKey = WScript.Arguments(1)
    sLoginServerName = WScript.Arguments(2)
End If
If WScript.Arguments.Count >= 4 Then
    sAuthType = WScript.Arguments(3)
Else
    sAuthType = "0"
End If
If WScript.Arguments.Count < 3 Then
    MsgBox "引数を正しく指定してください。", vbExclamation, WScript.ScriptName
    WScript.Quit
End If

Dim sMsgTitle
sMsgTitle = WScript.ScriptName

'=== 事前処理 ===
Dim sDownloadTrgtDirPathRaw
Dim sDownloadTrgtDirPath
Dim sScpProgramPath
Dim sDiffProgramPath
Dim sCodesDirPath
Dim sHomeDirPath
sDownloadTrgtDirPathRaw = objWshShell.SpecialFolders("Desktop")
sScpProgramPath = GetEnvVariable("MYEXEPATH_WINSCP")
sDiffProgramPath = GetEnvVariable("MYEXEPATH_WINMERGE")
sCodesDirPath = GetEnvVariable("MYDIRPATH_CODES")
sHomeDirPath = "/home/" & sLinuxUserName

Dim sLinuxCodesDirPath
sLinuxCodesDirPath = sCodesDirPath & "\linux"

'バックアップフォルダ作成
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix
Call CreateDirectry( sDownloadTrgtDirPath )

Dim vAnswer
Dim iIdx
vAnswer = MsgBox("ファイルを受信します。ダウンロードを開始しますか？", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    On Error Resume Next 'リモートにファイルが存在しなくても無視する
    '=== ファイル受信(Remote → Local) ===
    Dim sResultMsg
    sResultMsg = ""
    Dim sRelPathWin, sRelPathLinux, sRelSubDirBk
    For iIdx = 0 To csSYNC_TRGT_FILES.Count - 1
        sRelPathWin = Replace(csSYNC_TRGT_FILES(iIdx), "/", "\")
        sRelPathLinux = Replace(csSYNC_TRGT_FILES(iIdx), "\", "/")
        If InStrRev(sRelPathWin, "\") > 0 Then
            sRelSubDirBk = Left(sRelPathWin, InStrRev(sRelPathWin, "\") - 1)
            Call CreateDirectry(sDownloadTrgtDirPath & "\" & sRelSubDirBk)
        End If
        Dim sCmd
        If sAuthType = "0" Then ' Password
            sCmd = """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sLinuxUserName & ":" & sKey & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/" & sRelPathLinux & " " & sDownloadTrgtDirPath & "\" & sRelPathWin & """ ""exit"""
        Else                    ' PrivateKey
            sCmd = """" & sScpProgramPath & """ /console /privatekey=" & sKey & " /command ""option batch on"" ""open " & sLinuxUserName & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/" & sRelPathLinux & " " & sDownloadTrgtDirPath & "\" & sRelPathWin & """ ""exit"""
        End If
        objWshShell.Run sCmd, 0, True
        'MsgBox sCmd
        If objFSO.FileExists(sDownloadTrgtDirPath & "\" & sRelPathWin) Then
            'Do Nothing
        Else
            sResultMsg = sResultMsg & vbNewLine & csSYNC_TRGT_FILES(iIdx) & " のダウンロードはスキップされました。"
            'objFSO.CopyFile sLinuxCodesDirPath & "\" & sRelPathWin, sDownloadTrgtDirPath & "\" & sRelPathWin
        End If
    Next
    If sResultMsg <> "" Then
        MsgBox sResultMsg, vbOkOnly, sMsgTitle
    End If
    '=== ファイルバックアップ ===
    For iIdx = 0 To csSYNC_TRGT_FILES.Count - 1
        sRelPathWin = Replace(csSYNC_TRGT_FILES(iIdx), "/", "\")
        objFSO.CopyFile sDownloadTrgtDirPath & "\" & sRelPathWin, sDownloadTrgtDirPath & "\" & sRelPathWin & "_rmtorg"
    Next
    '=== フォルダ比較 ===
    For iIdx = 0 To csSYNC_TRGT_FILES.Count - 1
        sRelPathWin = Replace(csSYNC_TRGT_FILES(iIdx), "/", "\")
        objWshShell.Run """" & sDiffProgramPath & """ -r -s """ & sLinuxCodesDirPath & "\" & sRelPathWin & """ """ & sDownloadTrgtDirPath & "\" & sRelPathWin & """", 10, False
    Next
    On Error Goto 0
    MsgBox "比較/マージが完了したらOKを押してください。", vbOkOnly, sMsgTitle
Else
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sMsgTitle
    WScript.Quit
End If

vAnswer = MsgBox("編集したファイルを送信しますか？", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    '=== ファイル送信(Local → Remote) ===
    On Error Resume Next 'ローカルにファイルが存在しなくても無視する
    For iIdx = 0 To csSYNC_TRGT_FILES.Count - 1
        sRelPathWin = Replace(csSYNC_TRGT_FILES(iIdx), "/", "\")
        sRelPathLinux = Replace(csSYNC_TRGT_FILES(iIdx), "\", "/")
        If sAuthType = "0" Then ' Password
            sCmd = """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sLinuxUserName & ":" & sKey & "@" & sLoginServerName & """ ""put " & sDownloadTrgtDirPath & "\" & sRelPathWin & " " & sHomeDirPath & "/" & sRelPathLinux & """ ""exit"""
        Else                    ' PrivateKey
            sCmd = """" & sScpProgramPath & """ /console /privatekey=" & sKey & " /command ""option batch on"" ""open " & sLinuxUserName & "@" & sLoginServerName & """ ""put " & sDownloadTrgtDirPath & "\" & sRelPathWin & " " & sHomeDirPath & "/" & sRelPathLinux & """ ""exit"""
        End If
        objWshShell.Run sCmd, 0, True
        'MsgBox sCmd
    Next
    On Error Goto 0
End If

vAnswer = MsgBox("ダウンロードしたファイルを削除しますか？", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    '=== フォルダ削除 ===
    'objFSO.DeleteFolder sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix, True
    Call MoveToTrushBox(sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix)
End If

'MsgBox "処理が完了しました！", vbYesNo, sMsgTitle

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
        MsgBox "環境変数「" & sEnvVar & "」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sMsgTitle
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

