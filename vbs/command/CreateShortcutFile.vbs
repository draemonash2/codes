Option Explicit

'<usage>
'  CreateShortcutFile.vbs <shortcut_file_path> <shortcut_target_path> [<arguments> <workingdirectory> <hotkey> <windowstyle> <description>]
'    windowstyle : 4���ʏ�A3���ő剻�A7���ŏ���
'      �e�I�v�V�����̏ڍׂ� https://atmarkit.itmedia.co.jp/ait/articles/0712/27/news083_2.html �Q��
'  
'  ex)
'    CreateShortcutFile.vbs C:\Users\test.lnk C:\codes\vbs\test

Dim sShrtctFilePath
Dim sShrtctTrgtPath
Dim sArguments
Dim sHotKey
Dim sDescription
Dim sWindowStyle
Dim sWorkingDirectory
If WScript.Arguments.Count < 2 Then
	WScript.Echo "�w�肷�����������Ă��܂�:" & WScript.Arguments.Count
	WScript.Quit
End If
If WScript.Arguments.Count >= 2 Then
	sShrtctFilePath = WScript.Arguments(0)
	sShrtctTrgtPath = WScript.Arguments(1)
End If
If WScript.Arguments.Count >= 3 Then
	sArguments = WScript.Arguments(2)
End If
If WScript.Arguments.Count >= 4 Then
	sWorkingDirectory = WScript.Arguments(3)
End If
If WScript.Arguments.Count >= 5 Then
	sHotKey = WScript.Arguments(4)
End If
If WScript.Arguments.Count >= 6 Then
	sWindowStyle = WScript.Arguments(5)
End If
If WScript.Arguments.Count >= 7 Then
	sDescription = WScript.Arguments(6)
End If

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.CreateShortcut( sShrtctFilePath )
	.TargetPath = sShrtctTrgtPath
	.Arguments = sArguments
	.WorkingDirectory = sWorkingDirectory
	.HotKey = sHotKey
	.WindowStyle = CInt(sWindowStyle)
	.Description = sDescription
	.Save
End With
