'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ファイルパスをコピー"

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** 選択ファイル取得 ***
If bIsContinue = True Then
	If EXECUTION_MODE = 0 Then 'Explorerから実行
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		Dim sArg
		For Each sArg In WScript.Arguments
			cFilePaths.add sArg
		Next
	ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
		Set cFilePaths = WScript.Col( WScript.Env("Selected") )
	Else 'デバッグ実行
		MsgBox "デバッグモードです。"
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
	End If
Else
	'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
	If cFilePaths.Count = 0 Then
		MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
		MsgBox "処理を中断します", vbYes, PROG_NAME
		bIsContinue = False
	Else
		'Do Nothing
	End If
Else
	'Do Nothing
End If

'*** 相対パスに置換 ***
Dim objWshShellEnv
Set objWshShellEnv = WScript.CreateObject("WScript.Shell").Environment("Process")
sMatchDirName = objWshShellEnv.Item("MATCH_DIR_NAME")
lRemeveDirLevel = objWshShellEnv.Item("REMOVE_DIR_LEVEL")
Dim oFilePath
Dim sRelativePath
If sMatchDirName = "" And lRemeveDirLevel = "" Then
	'Do Nothing
Else
	Dim cRltvFilePaths
	Set cRltvFilePaths = CreateObject("System.Collections.ArrayList")
	For Each oFilePath In cFilePaths
		cRltvFilePaths.Add ReplaceRelatevePath(oFilePath, sMatchDirName, lRemeveDirLevel)
	Next
	Set cFilePaths = cRltvFilePaths
End If

'*** クリップボードへコピー ***
If bIsContinue = True Then
	Dim sOutString
	Dim bFirstStore
	bFirstStore = True
	For Each oFilePath In cFilePaths
		If bFirstStore = True Then
			sOutString = oFilePath
			bFirstStore = False
		Else
			sOutString = sOutString & vbNewLine & oFilePath
		End If
	Next
	CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
	'Do Nothing
End If

'相対パスへ置換
Private Function ReplaceRelatevePath( _
	ByVal sInFilePath, _
	ByVal sMatchDirName, _
	ByVal lRemeveDirLevel _
)
	Dim sRemoveDirLevelPath
	sRemoveDirLevelPath = ""
	Dim lIdx
	For lIdx = 0 To lRemeveDirLevel - 1
		sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
	Next
	
	Dim sSearchPattern
	Dim sTargetStr
	sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
	sTargetStr = sInFilePath
	
	Dim oRegExp
	Set oRegExp = CreateObject("VBScript.RegExp")
	oRegExp.Pattern = sSearchPattern				'検索パターンを設定
	oRegExp.IgnoreCase = True						'大文字と小文字を区別しない
	oRegExp.Global = True							'文字列全体を検索
	
	Dim oMatchResult
	Set oMatchResult = oRegExp.Execute(sTargetStr)	'パターンマッチ実行
	
	Dim sOutFilePath
	sOutFilePath = ""
	If oMatchResult.Count > 0 THen
		sOutFilePath = Replace( sInFilePath, oMatchResult.item(0), "" )
	Else
		sOutFilePath = sInFilePath
	End If
	
	ReplaceRelatevePath = sOutFilePath
End Function
