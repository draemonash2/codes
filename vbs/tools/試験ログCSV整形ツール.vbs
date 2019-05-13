Option Explicit

'===============================================================================
'【説明】
'	試験ログCSVを整形し、CANapeでインポートできる形式に変換する。
'	整形する内容は以下の通り。
'		- 「Datatype」列を付与
'			「DataType」は data_type_list.csv より取得する。
'		- RAM名を置換する（ex. ram[0]:1 → ram_0:1）
'	data_type_list.csv が存在しない場合は、すべて uint8 と解釈して変換する。
'
'【使用方法】
'	使用方法は２通り。
'		◆フォルダ配下の全csvすべてを置換したい場合
'			1. 本スクリプトを置換したいフォルダに移動する。
'			2. 本スクリプトを実行する。(ダブルクリック)
'		◆１ファイルのみ整形したい場合
'			1. 整形したい試験ログCSVを本スクリプトへdrag&dropする
'
'【覚え書き】
'	なし
'
'【改訂履歴】
'	1.0.0	2019/05/13	新規作成
'===============================================================================

'===============================================================================
'= インクルード
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\..\_lib\String.vbs" )		'GetFileExt()
Call Include( sMyDirPath & "\..\_lib\FileSystem.vbs" )	'GetFileList3()
Call Include( sMyDirPath & "\..\_lib\Collection.vbs" )	'ReadTxtFileToCollection()
														'WriteTxtFileFrCollection()

'===============================================================================
' 設定
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST CREATE_BACKUP_FILE = False
CONST DEFAULT_DATA_TYPE = "uint8"

'===============================================================================
' 本処理
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp"
Const DATATYPE_ROW_KEYWORD = "DataType"

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
If objFSO.FileExists(sDataTypeListFilePath) Then
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
		On Error Resume Next '重複キーがあったら無視
		oDataTypeList.Add objWords(0), objWords(1)
		On Error Goto 0
	Loop
	objTxtFile.Close
Else
	'Do Nothing
End If

'*****************************
' 試験ログCSVファイルリスト取得
'*****************************
dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")
If WScript.Arguments.Count = 0 Then
	dim cFileList
	Set cFileList = CreateObject("System.Collections.ArrayList")
	call GetFileList3(sRootDirPath, cFileList, 1)

	dim sFilePath
	for each sFilePath in cFileList
		if objFSO.GetExtensionName(sFilePath) = "csv" And _
		   objFSO.GetFileName(sFilePath) <> DATA_TYPE_LIST_FILE_NAME then
			cCsvFileList.add sFilePath
		end if
	next
	Set cFileList = Nothing
ElseIf WScript.Arguments.Count = 1 And _
	objFSO.FileExists(WScript.Arguments(0)) Then
	cCsvFileList.add WScript.Arguments(0)
Else
	WScript.Echo "引数エラー"
	WScript.Quit
End If

'*****************************
' 試験ログCSV整形
'*****************************
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	'*** 試験ログCSVバックアップ出力 ***
	
	If CREATE_BACKUP_FILE = True then
		Dim sCsvBakFilePath
		sCsvBakFilePath = sCsvFilePath & ".bak"
		If objFSO.FileExists(sCsvBakFilePath) Then
			'Do Nothing
		Else
			objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
		End If
	End If
	
	'*** 試験ログCSVオープン ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToCollection(sCsvFilePath, cFileContents)
	
	'*** 変数名置換 ***
	cFileContents(0) = ReplaceKeyword(cFileContents(0))
	
	'*** Datatype挿入 ***
	If InStr(cFileContents(1), DATATYPE_ROW_KEYWORD) Then
		'Do Nothing
	Else
		Dim vRamNames
		vRamNames = Split(cFileContents(0), ",")
		Dim sRamNameRaw
		Dim sRamNameRep
		Dim sDataTypeLine
		sDataTypeLine = DATATYPE_ROW_KEYWORD
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
					sDataTypeLine = sDataTypeLine & "," & DEFAULT_DATA_TYPE
				end if
			end if
		next
		cFileContents.Insert 1, sDataTypeLine
	End If
	
	'*** csv出力 ***
	call WriteTxtFileFrCollection(sCsvFilePath, cFileContents)
	
	Set cFileContents = Nothing
next

Set objFSO = Nothing
Set cCsvFileList = Nothing
set oDataTypeList = Nothing

MsgBox "試験ログCSV 整形完了!"

'===============================================================================
' 関数
'===============================================================================
Private Function ReplaceKeyword( _
	byval sTrgtWord _
)
	Dim sOutWord
	sOutWord = sTrgtWord
	sOutWord = Replace(sOutWord, "[", "_")
	sOutWord = Replace(sOutWord, "]", "")
	ReplaceKeyword = sOutWord
End Function

'===============================================================================
'= インクルード関数
'===============================================================================
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
