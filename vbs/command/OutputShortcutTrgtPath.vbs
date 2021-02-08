Option Explicit

'<usage>
'  OutputShortcutTrgtPath.vbs [-v] <shortcut_file_path>
'  
'  ex1) OutputShortcutTrgtPath.vbs C:\Users\test.txt.lnk
'         Å® C:\trgtfile.txt
'  ex2) OutputShortcutTrgtPath.vbs -v C:\Users\test.txt.lnk
'         Å® C:\Users\test.txt.lnk<TAB>C:\trgtfile.txt

Dim sShrtctFilePath
Dim sShrtctTrgtPath
Dim bVerboseMode
If WScript.Arguments.Count = 1 Then
	bVerboseMode = False
	sShrtctFilePath = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
	If WScript.Arguments(0) = "-v" Then
		bVerboseMode = True
		sShrtctFilePath = WScript.Arguments(1)
	Else
		WScript.Echo "[error] arguments = " & WScript.Arguments(0)
		WScript.Quit
	End If
Else
    WScript.Echo "[error] argment num = " & WScript.Arguments.Count
	WScript.Quit
End If

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
sShrtctTrgtPath = objWshShell.CreateShortcut( sShrtctFilePath ).TargetPath
If bVerboseMode = True Then
	WScript.Echo sShrtctFilePath & vbTab & sShrtctTrgtPath
Else
	WScript.Echo sShrtctTrgtPath
End If

