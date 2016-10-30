Option Explicit

Private Function RunasCheck()
	Dim flgRunasMode
	Dim objWMI, osInfo, flag, objShell, os
	Dim strArgs
	Dim args
	
	Set args = WScript.Arguments
	
	flgRunasMode = False
	strArgs = ""
	
	' ƒtƒ‰ƒO‚ÌŽæ“¾
	If args.Count > 0 Then
		If UCase(args.item(0)) = "/RUNAS" Then
			flgRunasMode = True
		End If
		strArgs = strArgs & " " & args.item(0)
	End If
	
	Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	flag = false
	For Each os in osInfo
		If Left(os.Version, 3) >= 6.0 Then
			flag = True
		End If
	Next
	
	Set objShell = CreateObject("Shell.Application")
	If flgRunasMode = False Then
		If flag = True Then
			objShell.ShellExecute _
			"wscript.exe", _
			"""" & WScript.ScriptFullName & """" & " /RUNAS " & strArgs, "", _
			"runas", _
			1
			Wscript.Quit
		End If
	End If
End Function
