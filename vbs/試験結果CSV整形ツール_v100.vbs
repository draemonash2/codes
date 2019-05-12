Option Explicit

'==============================================================================
' 設定
'==============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"

'==============================================================================
' 本処理
'==============================================================================
Const DATA_ROW_KEYWORD = "DataType"
Const RAMNAME_ROW_KEYWORD = "Data"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*****************************
' DataType一覧取得
'*****************************
dim oDataTypeList
set oDataTypeList = CreateObject("Scripting.Dictionary")

Dim sRootDirPath
sRootDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )

Dim sDataTypeListFilePath
sDataTypeListFilePath = sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME

dim objTxtFile
set objTxtFile = objFSO.OpenTextFile(sDataTypeListFilePath, 1, True)

dim objWords
Dim sTxtLine
Do Until objTxtFile.AtEndOfStream
	sTxtLine = objTxtFile.ReadLine
	objWords = split(sTxtLine, ",")
	if InStr(objWords(0), "[") Then
		objWords(0) = ReplaceKeyword(objWords(0))
	Else
		'Do Nothing
	end if
	oDataTypeList.Add objWords(0), objWords(1)
Loop
objTxtFile.Close

'*****************************
' 結果csvファイルリスト取得
'*****************************
dim cFileList
Set cFileList = CreateObject("System.Collections.ArrayList")
call GetFileList2(sRootDirPath, cFileList, 1)

dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")
dim sFilePath
for each sFilePath in cFileList
	if objFSO.GetExtensionName(sFilePath) = "csv" And _
	   objFSO.GetFileName(sFilePath) <> DATA_TYPE_LIST_FILE_NAME then
		cCsvFileList.add sFilePath
	end if
next

'*****************************
' 試験結果csv整形
'*****************************
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	'*** 結果csvバックアップ出力 ***
	Dim sCsvBakFilePath
	sCsvBakFilePath = sCsvFilePath & ".bak"
	If objFSO.FileExists(sCsvBakFilePath) Then
		'Do Nothing
	Else
		objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
	End If

	'*** 結果csvオープン ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToArray(sCsvFilePath, cFileContents)

	'*** 変数名置換 ***
	cFileContents(0) = ReplaceKeyword(cFileContents(0))

	'*** Datatype挿入 ***
	If InStr(cFileContents(1), DATA_ROW_KEYWORD) Then
		'Do Nothing
	Else
		Dim vRamNames
		vRamNames = Split(cFileContents(0), ",")
		Dim sRamNameRaw
		Dim sRamNameRep
		Dim sDataTypeLine
		sDataTypeLine = DATA_ROW_KEYWORD
		for each sRamNameRaw In vRamNames
			If sRamNameRaw = RAMNAME_ROW_KEYWORD Then
				'Do Nothing
			else
				sRamNameRep = sRamNameRaw
				sRamNameRep = ReplaceKeyword(sRamNameRep)
				if oDataTypeList.Exists(sRamNameRaw) then
					sDataTypeLine = sDataTypeLine & "," & oDataTypeList.Item(sRamNameRaw)
				elseif oDataTypeList.Exists(sRamNameRep) then
					sDataTypeLine = sDataTypeLine & "," & oDataTypeList.Item(sRamNameRep)
				else
					sDataTypeLine = sDataTypeLine & ",uint8"
				end if
			end if
		next
		cFileContents.Insert 1, sDataTypeLine
	End If

	'*** csv出力 ***
	call WriteTxtFileFrArray(sCsvFilePath, cFileContents)
next

MsgBox "試験結果CSV整形完了!"

'==============================================================================
' 関数
'==============================================================================
Private Function ReplaceKeyword( _
	byval sTrgtWord _
)
	Dim sOutWord
	sOutWord = sTrgtWord
	sOutWord = Replace(sOutWord, "[", "_")
	sOutWord = Replace(sOutWord, "]", "")
	ReplaceKeyword = sOutWord
End Function

'==============================================================================
' ライブラリ
'==============================================================================
' ==================================================================
' = 概要	ファイル/フォルダパス一覧を取得する
' = 引数	sTrgtDir		String		[in]	対象フォルダ
' = 引数	cFileList		Collections [out]	ファイル/フォルダパス一覧
' = 引数	lFileListType	Long		[in]	取得する一覧の形式
' =													0：両方
' =													1:ファイル
' =													2:フォルダ
' =													それ以外：格納しない
' = 戻値	なし
' = 覚書	・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' ==================================================================
Public Function GetFileList2( _
	ByVal sTrgtDir, _
	ByRef cFileList, _
	ByVal lFileListType _
)
	Dim objFSO	'FileSystemObjectの格納先
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	
	'Dir コマンド実行（出力結果を一時ファイルに格納）
	Dim sTmpFilePath
	Dim sExecCmd
	sTmpFilePath = WScript.CreateObject( "WScript.Shell" ).CurrentDirectory & "\Dir.tmp"
	Select Case lFileListType
		Case 0:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a > """ & sTmpFilePath & """"
		Case 1:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:a-d > """ & sTmpFilePath & """"
		Case 2:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:d > """ & sTmpFilePath & """"
		Case Else: sExecCmd = ""
	End Select
	With CreateObject("Wscript.Shell")
		.Run "cmd /c" & sExecCmd, 7, True
	End With
	
	Dim objFile
	Dim sTextAll
	On Error Resume Next
	If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
		If Err.Number = 0 Then
			sTextAll = objFile.ReadAll
			sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
			Dim vFileList
			vFileList = Split( sTextAll, vbNewLine )
			Dim sFilePath
			For Each sFilePath In vFileList
				cFileList.add sFilePath
			Next
			objFile.Close
		Else
			WScript.Echo "ファイルが開けません: " & Err.Description
		End If
		Set objFile = Nothing	'オブジェクトの破棄
	Else
		WScript.Echo "エラー " & Err.Description
	End If	
	objFSO.DeleteFile sTmpFilePath, True
	Set objFSO = Nothing	'オブジェクトの破棄
	On Error Goto 0
End Function
'	Call Test_GetFileList2()
	Private Sub Test_GetFileList2()
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Dim sCurDir
		sCurDir = objFSO.GetParentFolderName( WScript.ScriptFullName )
		
		msgbox sCurDir
		
		Dim cFileList
		Set cFileList = CreateObject("System.Collections.ArrayList")
		Call GetFileList2( sCurDir, cFileList, 1 )
		
		dim sFilePath
		dim sOutput
		sOutput = ""
		for each sFilePath in cFileList
			sOutput = sOutput & vbNewLine & sFilePath
		next
		MsgBox sOutput
	End Sub

' ==================================================================
' = 概要	指定されたファイルパスから拡張子を抽出する。
' = 引数	sFilePath	String	[in]  ファイルパス
' = 戻値				String		  拡張子
' = 覚書	・拡張子がない場合、空文字を返却する
' =			・ファイル名も指定可能
' ==================================================================
Public Function GetFileExt( _
	ByVal sFilePath _
)
	Dim sFileName
	sFileName = GetFileName(sFilePath)
	If InStr(sFileName, ".") > 0 Then
		GetFileExt = ExtractTailWord(sFileName, ".")
	Else
		GetFileExt = ""
	End If
End Function
	'Call Test_GetFileExt()
	Private Sub Test_GetFileExt()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")	  'txt
		Result = Result & vbNewLine & GetFileExt("c:\codes\test")		  '
		Result = Result & vbNewLine & GetFileExt("test.txt")			  'txt
		Result = Result & vbNewLine & GetFileExt("test")				  '
		Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
		Result = Result & vbNewLine & GetFileExt("test.aaa.txt")		  'txt
		MsgBox Result
	End Sub

' ==================================================================
' = 概要	テキストファイルの中身を配列に格納
' = 引数	sTrgtFilePath	String		[in]	ファイルパス
' = 引数	cFileContents	Collections [out]	ファイルの中身
' = 戻値	なし
' = 覚書	なし
' ==================================================================
Public Function ReadTxtFileToArray( _
	ByVal sTrgtFilePath, _
	ByRef cFileContents _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTxtFile
	Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
	
	Do Until objTxtFile.AtEndOfStream
		cFileContents.add objTxtFile.ReadLine
	Loop
	
	objTxtFile.Close
End Function
'	Call Test_OpenTxtFile2Array()
	Private Sub Test_OpenTxtFile2Array()
		Dim cFileList
		Set cFileList = CreateObject("System.Collections.ArrayList")
		sFilePath = "C:\codes\vbs\試験結果CSV整形ツール\data_type_list.csv"
		call ReadTxtFileToArray( sFilePath, cFileList )
		
		dim sFilePath
		dim sOutput
		sOutput = ""
		for each sFilePath in cFileList
			sOutput = sOutput & vbNewLine & sFilePath
		next
		MsgBox sOutput
	End Sub

' ==================================================================
' = 概要	配列の中身をテキストファイルに書き出し
' = 引数	sTrgtFilePath	String		[in]	ファイルパス
' = 引数	cFileContents	Collections [in]	ファイルの中身
' = 戻値	なし
' = 覚書	なし
' ==================================================================
Public Function WriteTxtFileFrArray( _
	ByVal sTrgtFilePath, _
	ByRef cFileContents _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTxtFile
	Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
	
	Dim sFileLine
	For Each sFileLine In cFileContents
		objTxtFile.WriteLine sFileLine
	Next
	
	objTxtFile.Close
End Function
'	Call Test_WriteTxtFileFrArray()
	Private Sub Test_WriteTxtFileFrArray()
		Dim cFileContents
		Set cFileContents = CreateObject("System.Collections.ArrayList")
		cFileContents.Add "a"
		cFileContents.Add "b"
		cFileContents.Insert 1, "c"
		DIm sTrgtFilePath
		sTrgtFilePath = "C:\codes\vbs\試験結果CSV整形ツール\Test.csv"
		call WriteTxtFileFrArray( sTrgtFilePath, cFileContents )
	End Sub
