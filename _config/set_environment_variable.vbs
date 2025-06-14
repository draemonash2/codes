Option Explicit

Call ExecRunas()

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim dEnvVars
Set dEnvVars = CreateObject("Scripting.Dictionary")

'▼▼▼ 設定ここから ▼▼▼
Const lEXEC_MODE = 1			'1:追加 2:削除
Const sENV_TARGET = "User"	'System:システム環境変数 User:ユーザ環境変数
																											' +------+------+------+------------+--------------+
																											' |  xf  |	ahk |  vim | codes(vbs) | updatecodes  |
With dEnvVars																								' +------+------+------+------------+--------------+
	.Add "MYDIRPATH_DESKTOP"		,"%USERPROFILE%\OneDrive\デスクトップ"									' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_DOCUMENTS"		,"%USERPROFILE%\OneDrive\Documents"										' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_PICTURES"		,"%USERPROFILE%\OneDrive\Pictures"										' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_CODES"			,"C:\codes"																' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_CODES_SAMPLE"	,"C:\codes_sample"														' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_GITHUB_IO"		,"C:\github_io"															' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_OTHER"			,"C:\other"																' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_PRG"			,"C:\prg"																' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_PRG_EXE"		,"C:\prg_exe"															' |  −  |	−	|  −  |	 −		|	   −	   |
	.Add "MYDIRPATH_CODES_CONFIG"	,"%MYDIRPATH_CODES%\_config"											' |  −  |	−	|  −  |	 ○		|	   −	   |
	.Add "MYEXEPATH_HIDEMARU"		,"%MYDIRPATH_PRG_EXE%\Hidemaru\Hidemaru.exe"							' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_WINMERGE"		,"%MYDIRPATH_PRG_EXE%\WinMerge\WinMergeU.exe"							' |  〇  |	−	|  −  |	 〇		|	   〇	   |
	.Add "MYEXEPATH_GVIM"			,"%MYDIRPATH_PRG_EXE%\Vim\gvim.exe"										' |  〇  |	〇	|  −  |	 〇		|	   −	   |
	.Add "MYEXEPATH_VSCODE"			,"%MYDIRPATH_PRG_EXE%\VSCode\Code.exe"									' |  〇  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_CURSOR"			,"%MYDIRPATH_PRG%\cursor\Cursor.exe"									' |  〇  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_TRESGREP"		,"%MYDIRPATH_PRG_EXE%\TresGrep\TresGrep.exe"							' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_EVERYTHING"		,"%MYDIRPATH_PRG_EXE%\Everything\Everything.exe"						' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_DISKINFO3"		,"%MYDIRPATH_PRG_EXE%\diskinfo64\DiskInfo3.exe"							' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_NEEVIEW"		,"%MYDIRPATH_PRG_EXE%\NeeView\NeeView.exe"								' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_MASSIGRA"		,"%MYDIRPATH_PRG_EXE%\MassiGra\MassiGra.exe"							' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_LINAME"			,"%MYDIRPATH_PRG_EXE%\LiName\LiName.exe"								' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_EXCEL"			,"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"		' |  〇  |	−	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_XF"				,"%MYDIRPATH_PRG_EXE%\X-Finder\XF.exe"									' |  −  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_CALC"			,"%MYDIRPATH_PRG_EXE%\cCalc\cCalc.exe"									' |  −  |	〇	|  −  |	 −		|	   −	   |
'	.Add "MYEXEPATH_CALC"			,"calc"																	' |  −  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_RAPTURE"		,"%MYDIRPATH_PRG_EXE%\Rapture\rapture.exe"								' |  −  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_ITHOUGHTS"		,"%MYDIRPATH_PRG_EXE%\iThoughts\iThoughts.exe"							' |  −  |	〇	|  −  |	 −		|	   −	   |
	.Add "MYEXEPATH_CTAGS"			,"%MYDIRPATH_PRG_EXE%\Ctags\ctags.exe"									' |  −  |	−	|  〇  |	 −		|	   −	   |
	.Add "MYEXEPATH_GTAGS"			,"%MYDIRPATH_PRG_EXE%\Gtags\bin\gtags.exe"								' |  −  |	−	|  〇  |	 −		|	   −	   |
	.Add "MYEXEPATH_7Z"				,"%MYDIRPATH_PRG_EXE%\7-ZipPortable\App\7-Zip64\7z.exe"					' |  −  |	−	|  −  |	 〇		|	   ○	   |
	.Add "MYEXEPATH_WINSCP"			,"%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe"								' |  −  |	−	|  −  |	 〇		|	   −	   |
																											' +------+------+------+------------+--------------+
End With
'▲▲▲ 設定ここまで ▲▲▲

With objWshShell.Environment(sENV_TARGET)
	Dim vKey
	If lEXEC_MODE = 1 Then
		For Each vKey In dEnvVars
			.Item(vKey) = dEnvVars.Item(vKey)
		Next
		Msgbox "環境変数を追加しました", vbOkOnly, WScript.ScriptName
	ElseIf lEXEC_MODE = 2 Then
		For Each vKey In dEnvVars
			.Remove(vKey)
		Next
		Msgbox "環境変数を削除しました", vbOkOnly, WScript.ScriptName
	Else
		Msgbox "lEXEC_MODEエラー！", vbCritical, WScript.ScriptName
	End If
End With

' ==================================================================
' = 概要	管理者権限で実行する
' = 引数	なし
' = 戻値	なし
' = 戻値				Boolean		[out]	実行結果
' = 覚書	自動的に引数に影響を及ぼすため、要注意
' = 依存	なし
' = 所属	Windows.vbs
' ==================================================================
Public Function ExecRunas()
	Dim oArgs
	Dim bIsRunas
	Dim sArgs
	
	bIsRunas = False
	sArgs = ""
	Set oArgs = WScript.Arguments
	
	' フラグの取得
	If oArgs.Count > 0 Then
		If UCase(oArgs.item(0)) = "/RUNAS" Then
			bIsRunas = True
		End If
		sArgs = sArgs & " " & oArgs.item(0)
	End If
	
	Dim bIsExecutableOs
	bIsExecutableOs = false
	Dim oOsInfos
	Set oOsInfos = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
	Dim oOs
	For Each oOs in oOsInfos
		If Left(oOs.Version, 3) >= 6.0 Then
			bIsExecutableOs = True
		End If
	Next
	
	Dim oWshShell
	Set oWshShell = CreateObject("Shell.Application")
	ExecRunas = False
	If bIsRunas = False Then
		If bIsExecutableOs = True Then
			oWshShell.ShellExecute _
			"wscript.exe", _
			"""" & WScript.ScriptFullName & """" & " /RUNAS " & sArgs, "", _
			"runas", _
			1
			ExecRunas = True
			Wscript.Quit
		End If
	End If
End Function
