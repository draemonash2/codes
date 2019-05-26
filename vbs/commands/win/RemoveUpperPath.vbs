Option Explicit

'==================================================
' = 例）
' =	  【入力】
' =		クリップボード値："C:\t_endo\900_ソースコード\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c"
' =		SOURCE_DIR_NAME："900_ソースコード"
' =		REMOVE_DIR_LEVEL：2
' =   【出力】
' =		クリップボード値："src_b\sid\sid_can_2E.c"
'==================================================

'==================================================
'= 設定
'==================================================
Const SOURCE_DIR_NAME = "codes"
Const REMOVE_DIR_LEVEL = 1

'==================================================
'= 本処理
'==================================================
Dim sInputPath
Dim objHTML
Set objHTML = CreateObject("htmlfile")
'sInputPath = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
sInputPath = "C:\codes\vbs\test.vbs"

'msgbox sInputPath

Dim sRemoveDirLevelPath
sRemoveDirLevelPath = ""
Dim lIdx
For lIdx = 0 To REMOVE_DIR_LEVEL - 1
	sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
Next

Dim sSearchPattern
Dim sTargetStr
sSearchPattern = ".*\\" & SOURCE_DIR_NAME & sRemoveDirLevelPath & "\\"
'Msgbox sSearchPattern
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

'msgbox sOutPath

'With CreateObject("Wscript.Shell").Exec("clip")
'  .StdIn.Write sOutPath
'End With
Wscript.StdOut.WriteLine sOutPath

test
