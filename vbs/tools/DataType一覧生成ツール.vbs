Option Explicit

'===============================================================================
'【説明】
'	試験ログCSVからDataType一覧を生成する
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
Call Include( sMyDirPath & "\..\_lib\FileSystem.vbs" )	'GetFileList3()
Call Include( sMyDirPath & "\..\_lib\Collection.vbs" )	'ReadTxtFileToCollection()
														'WriteTxtFileFrCollection()

'===============================================================================
' 設定
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"

'===============================================================================
' 本処理
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp"
Const DATATYPE_ROW_KEYWORD = "DataType"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*****************************
' 試験ログCSVファイルリスト取得
'*****************************
dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")
Dim sRootDirPath
sRootDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
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
' DataType取得
'*****************************
dim oDataTypeList
Set oDataTypeList = CreateObject("System.Collections.ArrayList")
dim oDataTypeListDupChk '重複チェック用
set oDataTypeListDupChk = CreateObject("Scripting.Dictionary")
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	'*** 試験ログCSVオープン ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToCollection(sCsvFilePath, cFileContents)
	
	'*** DataType取得 ***
	If InStr(cFileContents(1), DATATYPE_ROW_KEYWORD) Then
		Dim vRamNames
		Dim vDataTypes
		vRamNames = Split(cFileContents(0), ",")
		vDataTypes = Split(cFileContents(1), ",")
		Dim sRamNameRaw
		Dim sRamNameRep
		Dim lIdx
		lIdx = 0
		for each sRamNameRaw In vRamNames
			If sRamNameRaw = RAMNAME_ROW_KEYWORD Then
				'Do Nothing
			else
				sRamNameRep = sRamNameRaw
				sRamNameRep = ReplaceKeyword(sRamNameRep)
				Dim sDataTypeListLine
				sDataTypeListLine = sRamNameRep & "," & vDataTypes(lIdx)
				If Not oDataTypeListDupChk.Exists( sDataTypeListLine ) Then
					oDataTypeList.Add sDataTypeListLine
					oDataTypeListDupChk.Add sDataTypeListLine, ""
				end if
			end if
			lIdx = lIdx + 1
		next
	Else
		'Do Nothing
	End If
	Set cFileContents = Nothing
next

'*****************************
' DataType一覧出力
'*****************************
call WriteTxtFileFrCollection(sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME, oDataTypeList)

MsgBox "DataType一覧 生成完了!"

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set oDataTypeList = Nothing
set oDataTypeListDupChk = Nothing

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
