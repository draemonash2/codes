Option Explicit

'usage
' cscript.exe .\RemoveUpperPath.vbs <src_path> <dir_name> <remeve_dir_level>
'
'usage ex.
' cscript.exe .\RemoveUpperPath.vbs "C:\test\900_ソースコード\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c" "900_ソースコード" 2
'   →"src_b\sid\sid_can_2E.c"
' cscript.exe .\RemoveUpperPath.vbs "C:\test\900_ソースコード\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c" "900_ソースコード" 3 | clip
'   →"sid\sid_can_2E.c" をクリップボードにコピー

'==================================================
'= 設定
'==================================================
Const DEFAULT_MATCH_DIR_NAME = "codes"
Const DEFAULT_REMOVE_DIR_LEVEL = 2

'==================================================
'= 本処理
'==================================================
Dim sInputPath
Dim sMatchDirName
Dim lRemeveDirLevel
sMatchDirName = DEFAULT_MATCH_DIR_NAME
lRemeveDirLevel = DEFAULT_REMOVE_DIR_LEVEL
If WScript.Arguments.Count = 1 Then
	sInputPath = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
	sInputPath = WScript.Arguments(0)
	sMatchDirName = WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 3 Then
	sInputPath = WScript.Arguments(0)
	sMatchDirName = WScript.Arguments(1)
	lRemeveDirLevel = WScript.Arguments(2)
Else
	WScript.StdOut.WriteLine "指定する引数が異なります：" & WScript.Arguments.Count
	WScript.StdOut.WriteLine "処理を中断します"
	WScript.Quit
End If

Dim sRemoveDirLevelPath
sRemoveDirLevelPath = ""
Dim lIdx
For lIdx = 0 To lRemeveDirLevel - 1
	sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
Next

Dim sSearchPattern
Dim sTargetStr
sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
sTargetStr = sInputPath

Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
oRegExp.Pattern = sSearchPattern				'検索パターンを設定
oRegExp.IgnoreCase = True						'大文字と小文字を区別しない
oRegExp.Global = True							'文字列全体を検索

Dim oMatchResult
Set oMatchResult = oRegExp.Execute(sTargetStr)	'パターンマッチ実行

Dim sOutPath
sOutPath = ""
If oMatchResult.Count > 0 THen
	sOutPath = Replace( sInputPath, oMatchResult.item(0), "" )
Else
	sOutPath = sInputPath
End If

If sOutPath = "" Then
	'Do Nothing
Else
	Wscript.StdOut.WriteLine sOutPath
End If
