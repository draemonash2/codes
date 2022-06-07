Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )    'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ConvDate2String()

'===============================================================================
'= 本処理
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 3 Then
    MsgBox "引数を正しく指定してください。", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sDownloadUrl
Dim sLocalObjName
Dim sDiffTrgtDirPath
sDiffTrgtDirPath = WScript.Arguments(0)
sDownloadUrl = WScript.Arguments(1)
sLocalObjName = WScript.Arguments(2)

Dim sOutputMsg
sOutputMsg = WScript.ScriptName & " 「" & sLocalObjName & "」"

'=== 事前処理 ===
Dim sDownloadTrgtDirPath
Dim sDownloadTrgtFilePath
Dim sDiffSrcOrgDirPath
Dim sDiffSrcNewDirPath
Dim sUnzipProgramPath
Dim sDiffProgramPath
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLocalObjName & ".zip"
sDiffSrcOrgDirPath = sDownloadTrgtDirPath & "\" & sLocalObjName
sDiffSrcNewDirPath = sDownloadTrgtDirPath & "\" & sLocalObjName & "_" & sDateSuffix
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If InStr(sUnzipProgramPath, "%") > 0 then
    MsgBox "環境変数「MYEXEPATH_7Z」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "環境変数「MYEXEPATH_WINMERGE」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
end if

Dim vAnswer
vAnswer = MsgBox("ダウンロードを開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== ダウンロード ===
Call DownloadFile(sDownloadUrl, sDownloadTrgtFilePath)

'=== 解凍 ===
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """", 0, True
objFSO.MoveFolder sDiffSrcOrgDirPath, sDiffSrcNewDirPath

'=== フォルダ比較 ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcNewDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

'=== ZIP、フォルダ削除 ===
objFSO.DeleteFile sDownloadTrgtFilePath, True
vAnswer = MsgBox("ダウンロードフォルダを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFolder sDiffSrcNewDirPath, True
End If

MsgBox "処理が完了しました！", vbYesNo, sOutputMsg

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

