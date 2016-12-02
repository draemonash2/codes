Option Explicit

'<<概要>>
'  格納先を変更したプログラムの設定ファイルを、元の格納先に戻す。
'  その際、作成したシンボリックリンクを削除して、格納先変更前の
'  状態に復元する。
'
'<<引数>>
'  引数１：退避元ファイル/フォルダパス
'  引数２：退避先ファイル/フォルダパス
'（引数３：ログファイルパス）
'          ↑指定しない場合、ログメッセージを標準出力する。
'
'<<処理順>>
'  １．シンボリックリンク削除
'  ２．退避元フォルダへのショートカット削除
'  ３．ファイル/フォルダ移動（退避先⇒退避元）
'  ４．退避時に作成したフォルダを削除
'
'<<覚書>>
'  ・退避先ファイル/フォルダパスが存在しない場合、処理しない。
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
	WScript.Echo "#########################################################" & vbNewLine & _
				 "### result : [error  ] argument number error! arg num is " & WScript.Arguments.Count
	WScript.Quit
End If

Dim sExecResult
If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
	'Do Nothing
Else
	sExecResult = "#########################################################" & vbNewLine & _
				  "### result : [error  ] runas exec error!"
	If bIsLogValid = True Then
		objLogFile.WriteLine sExecResult
	Else
		WScript.Echo sExecResult
	End If
	WScript.Quit
End If

sExecResult = "#########################################################" & vbNewLine & _
			  "### src    : " & WScript.Arguments(ARG_IDX_SRCPATH) & vbNewLine & _
			  "### dst    : " & WScript.Arguments(ARG_IDX_DSTPATH)
If bIsLogValid = True Then
	objLogFile.WriteLine sExecResult
Else
	WScript.Echo sExecResult
End If

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_DSTPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	sExecResult = "### result : [error  ] destination path is missing!"
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

Dim sSrcPath
Dim sDstPath
Dim sDstParentDirPath
Dim sShortcutPath
sSrcPath = WScript.Arguments(ARG_IDX_SRCPATH)
sDstPath = WScript.Arguments(ARG_IDX_DSTPATH)
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )
sShortcutPath = sDstPath & "_linksrc.lnk"

If sFileType = "folder" Then
	If objFSO.FolderExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c rmdir /s /q """ & sSrcPath & """", 0, True
	If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
	objFSO.MoveFolder sDstPath, sSrcPath
Else
	If objFSO.FileExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c del /a /q """ & sSrcPath & """", 0, True
	If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
	objFSO.MoveFile sDstPath, sSrcPath
End If
Call DeleteEmptyFolder( sDstPath )
sExecResult = "### target : " & sFileType & vbNewLine & _
			  "### result : [success] setting files are restored!"
If bIsLogValid = True Then
	objLogFile.WriteLine sExecResult
Else
	WScript.Echo sExecResult
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

