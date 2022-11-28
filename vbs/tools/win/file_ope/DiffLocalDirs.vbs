Option Explicit

Dim sOutputMsg
sOutputMsg = WScript.ScriptName

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim cFilePathList
Set cFilePathList = CreateObject("System.Collections.ArrayList")

If WScript.Arguments.Count <> 1 Then
    MsgBox "引数を正しく指定してください。", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sTrgtDirPath
sTrgtDirPath = WScript.Arguments(0)

Dim sDiffProgramPath
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "環境変数「MYEXEPATH_WINMERGE」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, WScript.ScriptName
    WScript.Quit
End if

Dim vAnswer
vAnswer = MsgBox("ローカルフォルダとGithubのフォルダを比較します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

'フォルダ比較対象走査＆フォルダ比較実行
Dim vDirPath
Dim sDirNameRaw
Dim sDirNameBase
Dim sParDirPath
Dim sDiffSrcDirPath
Dim sDiffTrgtDirPath
dim oFolder
Set oFolder = objFSO.getFolder(sTrgtDirPath)
For Each vDirPath In oFolder.subfolders
    'MsgBox vDirPath
    sParDirPath = objFSO.GetParentFolderName( vDirPath )
    sDirNameRaw = objFSO.GetFileName( vDirPath )
    If InStr(sDirNameRaw, "_local") > 0 Then
        sDirNameBase = Replace(sDirNameRaw, "_local", "")
        sDiffSrcDirPath = sParDirPath & "\" & sDirNameBase
        sDiffTrgtDirPath = sParDirPath & "\" & sDirNameRaw
        If objFSO.FolderExists( sParDirPath & "\" & sDirNameBase ) Then
            objWshShell.Run """" & sDiffProgramPath & """ -r -s """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, False
        Else
            'Do Nothing
        End If
    Else
        'Do Nothing
    End If
Next
Dim oFile
For Each oFile In oFolder.Files
    Dim sFilePathDst
    Dim sFilePathSrc
    sFilePathDst = oFile.Path
    If InStr(sFilePathDst, "_local.") > 0 Then
        sFilePathSrc = Replace(sFilePathDst, "_local.", ".")
        'MsgBox """" & sFilePathSrc & """ """ & sFilePathDst & """"
        If sFilePathSrc <> sFilePathDst And objFSO.FileExists( sFilePathSrc ) Then
            'MsgBox """" & sFilePathSrc & """ """ & sFilePathDst & """"
            objWshShell.Run """" & sDiffProgramPath & """ -r -s """ & sFilePathSrc & """ """ & sFilePathDst & """", 10, False
        End If
    End If
Next

MsgBox "比較/マージが完了したらOKを押してください。", vbOkOnly, sOutputMsg

'MsgBox "完了！", vbOkOnly, WScript.ScriptName

