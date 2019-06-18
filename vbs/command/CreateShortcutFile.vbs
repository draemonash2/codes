Option Explicit

'<usage>
'  CreateShortcutFile.vbs <shortcut_file_path> <shortcut_target_path>
'  
'  ex)
'    CreateShortcutFile.vbs C:\Users\test.lnk C:\codes\vbs\test

Dim sShrtctFilePath
Dim sShrtctTrgtPath
If WScript.Arguments.Count = 2 Then
	sShrtctFilePath = WScript.Arguments(0)
	sShrtctTrgtPath = WScript.Arguments(1)
Else
	WScript.Echo "w’è‚·‚éˆø”‚ªŒë‚Á‚Ä‚¢‚Ü‚·:" & WScript.Arguments.Count
	WScript.Quit
End If

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.CreateShortcut( sShrtctFilePath )
	.TargetPath = sShrtctTrgtPath
	.Save
End With
