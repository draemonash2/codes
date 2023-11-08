Option Explicit

Call ExecRunas()

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim dEnvVars
Set dEnvVars = CreateObject("Scripting.Dictionary")

'������ �ݒ肱������ ������
Const lEXEC_MODE = 1			'1:�ǉ� 2:�폜
Const sENV_TARGET = "User"	'System:�V�X�e�����ϐ� User:���[�U���ϐ�
																											' +------+------+------+------------+--------------+
																											' |  xf  |	ahk |  vim | codes(vbs) | updatecodes  |
With dEnvVars																								' +------+------+------+------------+--------------+
	.Add "MYDIRPATH_DESKTOP"		,"%USERPROFILE%\OneDrive\�f�X�N�g�b�v"									' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_DOCUMENTS"		,"%USERPROFILE%\OneDrive\Documents"										' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_PICTURES"		,"%USERPROFILE%\OneDrive\Pictures"										' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_CODES"			,"C:\codes"																' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_CODES_SAMPLE"	,"C:\codes_sample"														' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_GITHUB_IO"		,"C:\github_io"															' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_OTHER"			,"C:\other"																' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_PRG"			,"C:\prg"																' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_PRG_EXE"		,"C:\prg_exe"															' |  �|  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYDIRPATH_CODES_CONFIG"	,"%MYDIRPATH_CODES%\_config"											' |  �|  |	�|	|  �|  |	 ��		|	   �|	   |
	.Add "MYEXEPATH_HIDEMARU"		,"%MYDIRPATH_PRG_EXE%\Hidemaru\Hidemaru.exe"							' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_WINMERGE"		,"%MYDIRPATH_PRG_EXE%\WinMerge\WinMergeU.exe"							' |  �Z  |	�|	|  �|  |	 �Z		|	   �Z	   |
	.Add "MYEXEPATH_GVIM"			,"%MYDIRPATH_PRG_EXE%\Vim\gvim.exe"										' |  �Z  |	�Z	|  �|  |	 �Z		|	   �|	   |
	.Add "MYEXEPATH_VSCODE"			,"%MYDIRPATH_PRG_EXE%\VSCode\Code.exe"									' |  �Z  |	�Z	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_TRESGREP"		,"%MYDIRPATH_PRG_EXE%\TresGrep\TresGrep.exe"							' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_EVERYTHING"		,"%MYDIRPATH_PRG_EXE%\Everything\Everything.exe"						' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_DISKINFO3"		,"%MYDIRPATH_PRG_EXE%\diskinfo64\DiskInfo3.exe"							' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_NEEVIEW"		,"%MYDIRPATH_PRG_EXE%\NeeView\NeeView.exe"								' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_MASSIGRA"		,"%MYDIRPATH_PRG_EXE%\MassiGra\MassiGra.exe"							' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_LINAME"			,"%MYDIRPATH_PRG_EXE%\LiName\LiName.exe"								' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_EXCEL"			,"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"		' |  �Z  |	�|	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_XF"				,"%MYDIRPATH_PRG_EXE%\X-Finder\XF.exe"									' |  �|  |	�Z	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_CALC"			,"%MYDIRPATH_PRG_EXE%\cCalc\cCalc.exe"									' |  �|  |	�Z	|  �|  |	 �|		|	   �|	   |
'	.Add "MYEXEPATH_CALC"			,"calc"																	' |  �|  |	�Z	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_RAPTURE"		,"%MYDIRPATH_PRG_EXE%\Rapture\rapture.exe"								' |  �|  |	�Z	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_ITHOUGHTS"		,"%MYDIRPATH_PRG_EXE%\iThoughts\iThoughts.exe"							' |  �|  |	�Z	|  �|  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_CTAGS"			,"%MYDIRPATH_PRG_EXE%\Ctags\ctags.exe"									' |  �|  |	�|	|  �Z  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_GTAGS"			,"%MYDIRPATH_PRG_EXE%\Gtags\bin\gtags.exe"								' |  �|  |	�|	|  �Z  |	 �|		|	   �|	   |
	.Add "MYEXEPATH_7Z"				,"%MYDIRPATH_PRG_EXE%\7-ZipPortable\App\7-Zip64\7z.exe"					' |  �|  |	�|	|  �|  |	 �Z		|	   ��	   |
	.Add "MYEXEPATH_WINSCP"			,"%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe"								' |  �|  |	�|	|  �|  |	 �Z		|	   �|	   |
																											' +------+------+------+------------+--------------+
End With
'������ �ݒ肱���܂� ������

With objWshShell.Environment(sENV_TARGET)
	Dim vKey
	If lEXEC_MODE = 1 Then
		For Each vKey In dEnvVars
			.Item(vKey) = dEnvVars.Item(vKey)
		Next
		Msgbox "���ϐ���ǉ����܂���", vbOkOnly, WScript.ScriptName
	ElseIf lEXEC_MODE = 2 Then
		For Each vKey In dEnvVars
			.Remove(vKey)
		Next
		Msgbox "���ϐ����폜���܂���", vbOkOnly, WScript.ScriptName
	Else
		Msgbox "lEXEC_MODE�G���[�I", vbCritical, WScript.ScriptName
	End If
End With

' ==================================================================
' = �T�v	�Ǘ��Ҍ����Ŏ��s����
' = ����	�Ȃ�
' = �ߒl	�Ȃ�
' = �ߒl				Boolean		[out]	���s����
' = �o��	�����I�Ɉ����ɉe�����y�ڂ����߁A�v����
' = �ˑ�	�Ȃ�
' = ����	Windows.vbs
' ==================================================================
Public Function ExecRunas()
	Dim oArgs
	Dim bIsRunas
	Dim sArgs
	
	bIsRunas = False
	sArgs = ""
	Set oArgs = WScript.Arguments
	
	' �t���O�̎擾
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
