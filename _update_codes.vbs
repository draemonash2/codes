Option Explicit

Const sDownloadUrl = "https://github.com/draemonash2/codes/archive/master.zip"
Const sDownloadTrgtFileName = "codes.zip"
Const sDiffSrcDirName = "codes-master"
Const sDiffTrgtDirPath = "C:\codes"
Const lPopupWaitSecond = 5

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sPopupTitle
sPopupTitle = WScript.ScriptName
Dim sDownloadTrgtDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
Dim sDownloadTrgtFilePath
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sDownloadTrgtFileName
Dim sDiffSrcDirPath
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sDiffSrcDirName

'=== ダウンロード ===
Dim sPopupMsg
sPopupMsg = "ダウンロードを開始します…"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation
Call DownloadFile( sDownloadUrl, sDownloadTrgtFilePath )

'=== 解凍 ===
sPopupMsg = "ダウンロード完了!" & vbNewLine & "解凍を開始します…"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation
With CreateObject("Shell.Application")
    .NameSpace(sDownloadTrgtDirPath).CopyHere .NameSpace(sDownloadTrgtFilePath).Items
End With

'=== 比較 ===
sPopupMsg = "解凍完了!" & vbNewLine & "比較を開始します…"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation

Dim sDiffProgramPath
sDiffProgramPath = objWshShell.Environment("System").Item("MYSYSPATH_WINMERGE")
If sDiffProgramPath = "" then
	MsgBox "環境変数が設定されていません。" & vbNewLine & "処理を中断します。", vbYes, PROG_NAME
	WScript.Quit
end if
objWshShell.Run sDiffProgramPath & " " & sDiffSrcDirPath & " " & sDiffTrgtDirPath, 0, True

'[参考URL] https://viewse.blogspot.com/2013/08/vbscriptweb.html
Private Function DownloadFile( _
    ByVal sDownloadUrl, _
    ByVal sDownloadTrgtFilePath _
)
    ' ダウンロード用のオブジェクト
    Dim objSrvHTTP
    Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")
    on error resume next
    Call objSrvHTTP.Open("GET", sDownloadUrl, False )
    if Err.Number <> 0 then
        Wscript.Echo Err.Description
        Wscript.Quit
    end if
    objSrvHTTP.Send
    
    if Err.Number <> 0 then
    ' おそらくサーバーの指定が間違っている
        Wscript.Echo Err.Description
        Wscript.Quit
    end if
    on error goto 0
    if objSrvHTTP.status = 404 then
        Wscript.Echo "URLが正しくありません(404)"
        Wscript.Quit
    end if
    
    ' バイナリデータ保存用オブジェクト
    Dim Stream
    Set Stream = Wscript.CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 ' バイナリ
    ' 戻されたバイナリをファイルとしてストリームに書き込み
    Stream.Write objSrvHTTP.responseBody
    ' ファイルとして保存
    Stream.SaveToFile sDownloadTrgtFilePath, 2
    Stream.Close
End Function

