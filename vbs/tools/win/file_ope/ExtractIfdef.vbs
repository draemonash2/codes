Option Explicit

'<概要>
'  ソースコード内の「#ifdef __ve__」をもとに、環境に合わせた処理を抽出して出力する。
'
'    ex. 「test.f」を本スクリプトにドラッグ＆ドロップした場合
'      "#ifdef __ve__" ～"#else" 部を抽出したソースコードを「test_ifside.f」として出力する。
'      "#else" ～"#endif" 部を抽出したソースコードを「test_elseside.f」として出力する。
'
'<使い方>
'  1. 除去したいソースコードファイルを本スクリプトにドラッグ＆ドロップする。

Const sSCRIPT_NAME = "IFDEF抽出ツール"

' ▼▼▼ 設定 ここから ▼▼▼
Const sKEYWORD_IF = "#ifdef __ve__"
Const sKEYWORD_ELSE = "#else /** __ve__ **/"
Const sKEYWORD_ENDIF = "#endif /** __ve__ **/"
Const sFILE_SUFFIX_IFSIDE = "ifside"
Const sFILE_SUFFIX_ELSESIDE = "elseside"
' ▲▲▲ 設定 ここまで ▲▲▲

If WScript.Arguments.Count = 0 Then
	MsgBox "最低一つは引数が必要です。処理を中断します。"
	WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sInputFilePath
Dim sOutIfSideFilePath
Dim sOutElseSideFilePath
sInputFilePath = WScript.Arguments(0)
sOutIfSideFilePath = _
	objFSO.GetParentFolderName( sInputFilePath ) & "\" & _
	objFSO.GetBaseName( sInputFilePath ) & _
	"_" & sFILE_SUFFIX_IFSIDE & _
	"." & objFSO.GetExtensionName( sInputFilePath )
sOutElseSideFilePath = _
	objFSO.GetParentFolderName( sInputFilePath ) & "\" & _
	objFSO.GetBaseName( sInputFilePath ) & _
	"_" & sFILE_SUFFIX_ELSESIDE & _
	"." & objFSO.GetExtensionName( sInputFilePath )

Dim adoInFile
Set adoInFile = CreateObject("ADODB.Stream")
Dim adoOutIfSideFile
Set adoOutIfSideFile = CreateObject("ADODB.Stream")
Dim adoOutElseSideFile
Set adoOutElseSideFile = CreateObject("ADODB.Stream")

adoInFile.Type = 2
adoInFile.Charset = "Shift_JIS"
adoInFile.LineSeparator = 10
adoInFile.Open
adoInFile.LoadFromFile sInputFilePath
adoOutIfSideFile.Type = 2
adoOutIfSideFile.Charset = "Shift_JIS"
adoOutIfSideFile.LineSeparator = 10
adoOutIfSideFile.Open
adoOutElseSideFile.Type = 2
adoOutElseSideFile.Charset = "Shift_JIS"
adoOutElseSideFile.LineSeparator = 10
adoOutElseSideFile.Open

Const lMODE_DEFAULT = 0
Const lMODE_IFDEF = 1
Const lMODE_ELSE = 2

Dim lLineNo
Dim lKeywordMode
Dim bIsError
Dim sErrMsg
lLineNo = 1
lKeywordMode = lMODE_DEFAULT
sErrMsg = ""

Do Until adoInFile.EOS
	Dim sLine
    sLine = adoInFile.ReadText(-2)
	If InStr(sLine, sKEYWORD_IF) > 0 Then
		If lKeywordMode = lMODE_IFDEF Then
			sErrMsg = sErrMsg & "#ifdef後に#ifdefが見つかりました : "& lLineNo & "行目" & vbNewLine
		ElseIf lKeywordMode = lMODE_ELSE Then
			sErrMsg = sErrMsg & "#else後に#ifdefが見つかりました : "& lLineNo & "行目" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_IFDEF
	ElseIf InStr(sLine, sKEYWORD_ELSE) > 0 Then
		If lKeywordMode = lMODE_ELSE Then
			sErrMsg = sErrMsg & "#else後に#elseが見つかりました : "& lLineNo & "行目" & vbNewLine
		ElseIf lKeywordMode = lMODE_DEFAULT Then
			sErrMsg = sErrMsg & "#endif後に#elseが見つかりました : "& lLineNo & "行目" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_ELSE
	ElseIf InStr(sLine, sKEYWORD_ENDIF) > 0 Then
		If lKeywordMode = lMODE_DEFAULT Then
			sErrMsg = sErrMsg & "#endif後に#endifが見つかりました : "& lLineNo & "行目" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_DEFAULT
	Else
		'行出力
		If lKeywordMode = lMODE_IFDEF Then
			adoOutIfSideFile.WriteText sLine, 1
		ElseIf lKeywordMode = lMODE_ELSE Then
			adoOutElseSideFile.WriteText sLine, 1
		ElseIf lKeywordMode = lMODE_DEFAULT Then
			adoOutIfSideFile.WriteText sLine, 1
			adoOutElseSideFile.WriteText sLine, 1
		Else
			MsgBox "予期しないエラー！ lKeywordMode = " & lKeywordMode
		End If
	End If
	lLineNo = lLineNo + 1
Loop

If sErrMsg <> "" Then
	MsgBox _
		"エラーが見つかりました。" & vbNewLine & _
		"処理を中断します。" & vbNewLine & _
		"---" & vbNewLine & _
		sErrMsg, _
		vbOkOnly, _
		sSCRIPT_NAME
Else
	adoOutIfSideFile.SaveToFile sOutIfSideFilePath, 2
	adoOutElseSideFile.SaveToFile sOutElseSideFilePath, 2
	MsgBox "完了！", vbOkOnly, sSCRIPT_NAME
End If

adoInFile.Close
adoOutIfSideFile.Close
adoOutElseSideFile.Close

