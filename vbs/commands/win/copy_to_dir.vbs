Option Explicit

'==============================================================================
'【説明】
'	ファイル/フォルダをコピーする。
'	移動先のフォルダが存在しない場合、フォルダを作成してからコピーする。
'
'【使用方法】
'	copy_to_dir.vbs <source_path> <destination_path>
'
'【使用例】
'	1) copy_to_dir.vbs c:\codes\vbs\test.txt c:\test\test.txt
'	2) copy_to_dir.vbs c:\codes\vbs c:\test\vbs
'		c:\codes\vbs
'			└ a.txt
'			└ b
'				└ c.txt
'		↓
'		c:\test\vbs
'			└ a.txt
'			└ b
'				└ c.txt
'
'【覚え書き】
'	なし
'
'【改訂履歴】
'	1.0.0	2019/05/12	新規作成
'==============================================================================

'==============================================================================
' 設定
'==============================================================================

'==============================================================================
'= インクルード
'==============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\String.vbs" )			'GetDirPath()
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )		'CreateDirectry()
														'GetFileOrFolder()

'==============================================================================
' 本処理
'==============================================================================
'引数チェック
If WScript.Arguments.Count = 2 Then
	'Do Nothing
Else
	Wscript.quit
End If

dim sSrcPath
dim sDstPath
sSrcPath = Replace(WScript.Arguments(0), "/", "\")
sDstPath = Replace(WScript.Arguments(1), "/", "\")

Dim lSrcPathType
lSrcPathType = GetFileOrFolder(sSrcPath)

dim sDstParDir
sDstParDir = GetDirPath( sDstPath )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If lSrcPathType = 1 Then 'ファイル
	call CreateDirectry( sDstParDir )
	objFSO.CopyFile sSrcPath, sDstPath
ElseIf lSrcPathType = 2 Then 'フォルダ
	call CreateDirectry( sDstParDir )
	objFSO.CopyFolder sSrcPath, sDstPath
Else '未存在
'	WScript.Echo "ファイルが存在しません"
End If

Set objFSO = Nothing

'==============================================================================
'= インクルード関数
'==============================================================================
' 外部プログラム インクルード関数
Private Function Include( _
	ByVal sOpenFile _
)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sOpenFile = objFSO.GetAbsolutePathName( sOpenFile )
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function
