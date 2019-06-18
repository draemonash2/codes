Option Explicit

'<usage>
'  OutputShortcutTrgtPath.vbs <shortcut_file_path>
'  
'  ex)
'    OutputShortcutTrgtPath.vbs C:\Users\test.lnk

Dim sShrtctFilePath
Dim sShrtctTrgtPath
If WScript.Arguments.Count = 1 Then
	sShrtctFilePath = WScript.Arguments(0)
Else
	WScript.Echo "w’è‚·‚éˆø”‚ªŒë‚Á‚Ä‚¢‚Ü‚·:" & WScript.Arguments.Count
	WScript.Quit
End If

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
sShrtctTrgtPath = objWshShell.CreateShortcut( sShrtctFilePath ).TargetPath
WScript.Echo sShrtctFilePath & vbTab & sShrtctTrgtPath
