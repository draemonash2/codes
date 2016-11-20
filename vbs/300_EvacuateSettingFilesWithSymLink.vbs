Option Explicit

'<<概要>>
'  プログラムの設定ファイルの格納先を変更する。
'  なお、元格納先から新格納先に向けてシンボリックリンクを作成するため、
'  プログラム側での設定は不要。
'
'<<引数>>
'  引数１：退避元ファイル/フォルダパス
'  引数２：退避先ファイル/フォルダパス
'（引数３：ログファイルパス）
'          ↑指定しない場合、ログメッセージを標準出力する。
'
'<<処理順>>
'  １．退避先ファイル/フォルダ削除
'  ２．退避先ファイル/フォルダ作成
'  ３．ファイル/フォルダ移動（退避元⇒退避先）
'  ４．退避元⇒退避先へのシンボリックリンク作成
'  ５．退避元フォルダへのショートカットを作成
'
'<<覚書>>
'  ・すでにシンボリックリンクが作成されている場合は、処理しない。
'  ・退避元ファイル/フォルダパスが存在しない場合、処理しない。
'  ・指定するパスはファイル/フォルダどちらでも可。
'  ・本スクリプト内で強制的に管理者権限に変更するため、
'    ローカル権限でも実行できる。ただし、本スクリプト呼び出し毎に
'    管理者権限実行の確認ウィンドウが表示されるため、呼び出し元で
'    あらかじめ管理者権限で実行しておくことをお勧めする。

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\lib\Windows.vbs" )
Call Include( sMyDirPath & "\lib\String.vbs" )

'==========================================================
'= 本処理
'==========================================================
Const ARG_COUNT_LOGVALID = 4
Const ARG_COUNT_LOGINVALID = 3
Const ARG_IDX_RUNAS = 0
Const ARG_IDX_SRCPATH = 1
Const ARG_IDX_DSTPATH = 2
Const ARG_IDX_LOGDIR = 3

'本スクリプトを管理者として実行させる
If ExecRunas( False ) Then WScript.Quit
	
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'###############################################
'# 事前処理
'###############################################
Dim bIsLogValid
If WScript.Arguments.Count = ARG_COUNT_LOGVALID Then
	bIsLogValid = True
	Dim objLogFile
	Set objLogFile = objFSO.OpenTextFile( WScript.Arguments(ARG_IDX_LOGDIR), 8, True) '第二引数：IOモード（1:読出し、2:新規書込み、8:追加書込み）
ElseIf WScript.Arguments.Count = ARG_COUNT_LOGINVALID Then
	bIsLogValid = False
Else
	WScript.Echo "[error] argument number error!" & vbNewLine & _
		   "  argument num : " & WScript.Arguments.Count
	WScript.Quit
End If

Dim sExecResult
If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
	'Do Nothing
Else
	sExecResult = "[error] runas exec error!"
	If bIsLogValid = True Then
		objLogFile.WriteLine sExecResult
	Else
		WScript.Echo sExecResult
	End If
	WScript.Quit
End If

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_SRCPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	sExecResult = "[error] source path is missing!" & vbNewLine & _
				  "  src : " & WScript.Arguments(ARG_IDX_SRCPATH) & vbNewLine & _
				  "  dst : " & WScript.Arguments(ARG_IDX_DSTPATH)
	If bIsLogValid = True Then
		objLogFile.WriteLine sExecResult
	Else
		WScript.Echo sExecResult
	End If
	WScript.Quit
End If

'###############################################
'# 本処理
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath

If sFileType = "folder" Then
	Dim sSrcDirPath
	Dim sDstDirPath
	Dim sSrcDirParentDirPath
	Dim sDstDirParentDirPath
	sSrcDirPath    = WScript.Arguments(ARG_IDX_SRCPATH)
	sDstDirPath    = WScript.Arguments(ARG_IDX_DSTPATH)
	sSrcDirParentDirPath = objFSO.GetParentFolderName( sSrcDirPath )
	sDstDirParentDirPath = objFSO.GetParentFolderName( sDstDirPath )
	If objFSO.GetFolder( sSrcDirPath ).Attributes And 1024 Then
		sExecResult = "[error] setting files are already evacuated!" & vbNewLine & _
					  "  src : " & sSrcDirPath & vbNewLine & _
					  "  dst : " & sDstDirPath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	Else
		If objFSO.FolderExists( sDstDirPath ) Then objFSO.DeleteFolder sDstDirPath, True
		Call CreateDirectry( GetDirPath( sDstDirPath ) )
		objFSO.MoveFolder sSrcDirPath, sDstDirPath
		objWshShell.Run "%ComSpec% /c mklink /d """ & sSrcDirPath & """ """ & sDstDirPath & """", 0, True
		sShortcutPath = sDstDirParentDirPath & "\" & GetFileName( sSrcDirPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcDirParentDirPath
				.Save
			End With
		End If
		sExecResult = "[success] setting files are evacuated!" & vbNewLine & _
					  "  src : " & sSrcDirPath & vbNewLine & _
					  "  dst : " & sDstDirPath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	End If
Else
	Dim sSrcFilePath
	Dim sDstFilePath
	Dim sDstFileParentDirPath
	Dim sSrcFileParentDirPath
	
	sSrcFilePath	= WScript.Arguments(ARG_IDX_SRCPATH)
	sDstFilePath	= WScript.Arguments(ARG_IDX_DSTPATH)
	sDstFileParentDirPath = objFSO.GetParentFolderName( sDstFilePath )
	sSrcFileParentDirPath = objFSO.GetParentFolderName( sSrcFilePath )
	
	If objFSO.GetFile( sSrcFilePath ).Attributes And 1024 Then
		sExecResult = "[error] setting files are already evacuated!" & vbNewLine & _
					  "  src : " & sSrcFilePath & vbNewLine & _
					  "  dst : " & sDstFilePath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	Else
		If objFSO.FileExists( sDstFilePath ) Then objFSO.DeleteFile sDstFilePath, True
		Call CreateDirectry( GetDirPath( sDstFilePath ) )
		objFSO.MoveFile sSrcFilePath, sDstFilePath
		objWshShell.Run "%ComSpec% /c mklink """ & sSrcFilePath & """ """ & sDstFilePath & """", 0, True
		sShortcutPath = sDstFileParentDirPath & "\" & GetFileName( sSrcFilePath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcFileParentDirPath
				.Save
			End With
		End If
		sExecResult = "[success] setting files are evacuated!" & vbNewLine & _
					  "  src : " & sSrcFilePath & vbNewLine & _
					  "  dst : " & sDstFilePath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	End If
End If

If bIsLogValid = True Then
	objLogFile.Close
	Set objLogFile = Nothing
Else
	'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing

'==========================================================
'= 関数定義
'==========================================================
' 外部プログラム インクルード関数
Function Include( _
	ByVal sOpenFile _
)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function

