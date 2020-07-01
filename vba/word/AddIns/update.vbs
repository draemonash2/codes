Const sTEMPLATE_FILE_NAME = "Normal.dotm"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sUserDirPath
sUserDirPath = CreateObject("Shell.Application").Namespace(40).Self.Path
Dim sDstTmpFilePath
Dim sSrcTmpFilePath
sDstTmpFilePath = sUserDirPath & "\AppData\Roaming\Microsoft\Templates\" & sTEMPLATE_FILE_NAME
sSrcTmpFilePath = objFSO.GetParentFolderName( WScript.ScriptFullName ) & "\" & sTEMPLATE_FILE_NAME

'Msgbox sDstTmpFilePath & vbNewLine & sSrcTmpFilePath
objFSO.CopyFile sSrcTmpFilePath, sDstTmpFilePath, True

Msgbox sTEMPLATE_FILE_NAME & "ÇçXêVÇµÇ‹ÇµÇΩÅI"

