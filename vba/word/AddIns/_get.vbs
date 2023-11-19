Const sTEMPLATE_FILE_NAME = "Normal.dotm"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sUserDirPath
sUserDirPath = CreateObject("Shell.Application").Namespace(40).Self.Path
Dim sDstTmpFilePath
Dim sSrcTmpFilePath
sSrcTmpFilePath = sUserDirPath & "\AppData\Roaming\Microsoft\Templates\" & sTEMPLATE_FILE_NAME
sDstTmpFilePath = objFSO.GetParentFolderName( WScript.ScriptFullName ) & "\" & sTEMPLATE_FILE_NAME

'Msgbox sDstTmpFilePath & vbNewLine & sSrcTmpFilePath
objFSO.CopyFile sSrcTmpFilePath, sDstTmpFilePath, True

Msgbox sTEMPLATE_FILE_NAME & "ÇéÊìæÇµÇ‹ÇµÇΩÅI"

