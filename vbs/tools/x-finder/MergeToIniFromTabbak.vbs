Option Explicit

Const sTRGT_FILE_NAME = "XF.ini"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sCurDirPath
sCurDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )

'=== バックアップ ===
objFSO.CopyFile sCurDirPath & "\" & sTRGT_FILE_NAME, sCurDirPath & "\" & sTRGT_FILE_NAME & ".bak", True

'=== .ini.tabbak からタブ情報取得 ===
Dim objTabBakFile
Set objTabBakFile = objFSO.OpenTextFile(sCurDirPath & "\" & sTRGT_FILE_NAME & ".tabbak", 1, True)
Dim vFileTabBak
vFileTabBak = Split(objTabBakFile.ReadAll, vbNewLine)
Dim sNewTabLines
Dim sTabSection ' blank or "Tab1" or "Tab2"
Dim sLine
sNewTabLines = ""
sTabSection = ""
For Each sLine In vFileTabBak
	If sLine = "[Tab]" Then
		sTabSection = "Tab1"
	ElseIf sLine = "[Tab2]" Then
		sTabSection = "Tab2"
	ElseIf sLine = "" Then
		If sTabSection = "Tab2" Then
			sTabSection = ""
		End If
	End If
	
	If sTabSection <> "" Then
		If sNewTabLines = "" Then
			sNewTabLines = sLine
		Else
			sNewTabLines = sNewTabLines & vbNewLine & sLine
		End If
	End If
Next
objTabBakFile.Close
'MsgBox "===" & vbNewLine & sNewTabLines & vbNewLine & "==="
'WScript.Quit

'=== マージ .ini.tabbak→.ini ===
Dim objOrgFile
Set objOrgFile = objFSO.OpenTextFile(sCurDirPath & "\" & sTRGT_FILE_NAME & ".bak", 1, True)

Dim vFileOrg
vFileOrg = Split(objOrgFile.ReadAll, vbNewLine)

sTabSection = ""
Dim sNewIniFileLines
Dim bIsTabBakMerged
sNewIniFileLines = ""
bIsTabBakMerged = False
For Each sLine In vFileOrg
	If sLine = "[Tab]" Then
		sTabSection = "Tab1"
	ElseIf sLine = "[Tab2]" Then
		sTabSection = "Tab2"
	ElseIf sLine = "" Then
		If sTabSection = "Tab2" Then
			sTabSection = ""
		End If
	End If
	
	If sTabSection = "" Then
		If sNewIniFileLines = "" Then
			sNewIniFileLines = sLine
		Else
			sNewIniFileLines = sNewIniFileLines & vbNewLine & sLine
		End If
	Else
		If bIsTabBakMerged = False Then
			If sNewIniFileLines = "" Then
				sNewIniFileLines = sNewTabLines
			Else
				sNewIniFileLines = sNewIniFileLines & vbNewLine & sNewTabLines
			End If
			bIsTabBakMerged = True
		End If
	End If
Next
objOrgFile.Close
'Dim vFileDummy
'vFileDummy = Split(sNewIniFileLines, vbNewLine)
'MsgBox "===" & vbNewLine & vFileDummy(UBound(vFileDummy)) & vbNewLine & "==="
'WScript.Quit

'=== ファイル出力 ===
Dim objNewFile
Set objNewFile = objFSO.OpenTextFile(sCurDirPath & "\" & sTRGT_FILE_NAME, 2, True)
objNewFile.Write sNewIniFileLines
objNewFile.Close

MsgBox "完了！"

