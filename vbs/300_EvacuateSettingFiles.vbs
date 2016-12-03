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
Call Include( sMyDirPath & "\lib\Log.vbs" )

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
Dim oLog
Set oLog = New LogMng
If WScript.Arguments.Count = ARG_COUNT_LOGVALID Then
	Call oLog.LogFileOpen( _
		WScript.Arguments(ARG_IDX_LOGDIR), _
		"+w" _
	)
ElseIf WScript.Arguments.Count = ARG_COUNT_LOGINVALID Then
	'Do Nothing
Else
	oLog.LogPuts "#########################################################"
	oLog.LogPuts "### result : [error  ] argument number error! arg num is " & WScript.Arguments.Count
	WScript.Quit
End If

If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
	'Do Nothing
Else
	oLog.LogPuts "#########################################################"
	oLog.LogPuts "### result : [error  ] runas exec error!"
	WScript.Quit
End If

oLog.LogPuts "#########################################################"
oLog.LogPuts "### src    : " & WScript.Arguments(ARG_IDX_SRCPATH)
oLog.LogPuts "### dst    : " & WScript.Arguments(ARG_IDX_DSTPATH)

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_SRCPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	oLog.LogPuts "### result : [error  ] source path is missing!"
	WScript.Quit
End If

'###############################################
'# 本処理
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath
Dim sSrcPath
Dim sDstPath
Dim sSrcParentDirPath
Dim sDstParentDirPath
sSrcPath    = WScript.Arguments(ARG_IDX_SRCPATH)
sDstPath    = WScript.Arguments(ARG_IDX_DSTPATH)
sSrcParentDirPath = objFSO.GetParentFolderName( sSrcPath )
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )

If sFileType = "folder" Then
	If objFSO.GetFolder( sSrcPath ).Attributes And 1024 Then
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [error  ] setting files are already evacuated!"
	Else
		If objFSO.FolderExists( sDstPath ) Then objFSO.DeleteFolder sDstPath, True
		Call CreateDirectry( GetDirPath( sDstPath ) )
		objFSO.MoveFolder sSrcPath, sDstPath
		objWshShell.Run "%ComSpec% /c mklink /d """ & sSrcPath & """ """ & sDstPath & """", 0, True
		sShortcutPath = sDstParentDirPath & "\" & GetFileName( sSrcPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcParentDirPath
				.Save
			End With
		End If
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [success] setting files are evacuated!"
	End If
Else
	If objFSO.GetFile( sSrcPath ).Attributes And 1024 Then
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [error  ] setting files are already evacuated!"
	Else
		If objFSO.FileExists( sDstPath ) Then objFSO.DeleteFile sDstPath, True
		Call CreateDirectry( GetDirPath( sDstPath ) )
		objFSO.MoveFile sSrcPath, sDstPath
		objWshShell.Run "%ComSpec% /c mklink """ & sSrcPath & """ """ & sDstPath & """", 0, True
		sShortcutPath = sDstParentDirPath & "\" & GetFileName( sSrcPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcParentDirPath
				.Save
			End With
		End If
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [success] setting files are evacuated!"
	End If
End If

Call oLog.LogFileClose

Set oLog = Nothing
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
