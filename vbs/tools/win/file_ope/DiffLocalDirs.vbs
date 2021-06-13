Option Explicit

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

'フォルダ比較対象走査＆フォルダ比較実行
Dim vDirPath
Dim sDirNameRaw
Dim sDirNameBase
Dim sParDirPath
Dim sDiffSrcDirPath
Dim sDiffTrgtDirPath
dim oFolder
set oFolder = objFSO.getFolder(sTrgtDirPath)
For each vDirPath in oFolder.subfolders
    'MsgBox vDirPath
    sParDirPath = objFSO.GetParentFolderName( vDirPath )
    sDirNameRaw = objFSO.GetFileName( vDirPath )
    If InStr(sDirNameRaw, "_local") > 0 Then
        sDirNameBase = Replace(sDirNameRaw, "_local", "")
        sDiffSrcDirPath = sParDirPath & "\" & sDirNameBase
        sDiffTrgtDirPath = sParDirPath & "\" & sDirNameRaw
        If objFSO.FolderExists( sParDirPath & "\" & sDirNameBase ) Then
            objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, False
        Else
            'Do Nothing
        End If
    Else
        'Do Nothing
    End If
Next

'MsgBox "完了！", vbOkOnly, WScript.ScriptName

