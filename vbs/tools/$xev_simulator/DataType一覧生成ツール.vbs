Option Explicit

'===============================================================================
'【概要】
'   試験ログCSVから変数シンボル名とDataTypeを抽出して「data_type_list.csv」を生成する
'
'【使用方法】
'   使用方法は２通り。
'       ◆フォルダ配下の全試験ログ(CSV)からDataTypeを抽出したい場合
'           １．抽出対象の試験ログ(CSV)と同階層以上のフォルダに
'               「DataType一覧生成ツール.vbs」を格納。
'           ２．「DataType一覧生成ツール.vbs」を実行する。
'       ◆１ファイルからのみ抽出したい場合
'           １．抽出したい試験ログ(CSV)を「DataType一覧生成ツール.vbs」へdrag&dropする。
'
'【詳細仕様】
'   ・ファイルの先頭に"TimeStamp"と記載された.csvファイルを試験ログ(CSV)と解釈する。
'   ・以下のような追加設定が可能。
'     - 変数シンボル名から配列識別子を除去する機能の有効無効
'         → REPLACE_RAM_SYMBOL = True:有効 / False:無効
'     - 整形完了時のメッセージ出力有無
'         → OUTPUT_FINISH_MESSAGE = True:出力 / False:出力しない
'
'【改訂履歴】
'   1.0.0   2019/07/01  遠藤    ・新規作成
'===============================================================================

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )          'GetFileList3()
                                                            'GetFileNotExistPath()
Call Include( "C:\codes\vbs\_lib\Collection.vbs" )          'ReadTxtFileToCollection()
                                                            'WriteTxtFileFrCollection()
Call Include( "C:\codes\vbs\_lib\ProgressBarCscript.vbs" )  'Class ProgressBar

'===============================================================================
' 設定
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST REPLACE_RAM_SYMBOL = False
CONST OUTPUT_FINISH_MESSAGE = True

'===============================================================================
' 本処理
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp,"
Const DATATYPE_ROW_KEYWORD = "DataType"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objPrgrsBar
Set objPrgrsBar = New ProgressBar
objPrgrsBar.Message = "DataType一覧生成中..."

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
    WScript.Echo "指定する引数の数が誤っています:" & WScript.Arguments.Count
    WScript.Quit
End If

'*****************************
' DataType取得
'*****************************
dim cDataTypeList
Set cDataTypeList = CreateObject("System.Collections.ArrayList")
dim dDataTypeListDupChk '重複チェック用
set dDataTypeListDupChk = CreateObject("Scripting.Dictionary")
dim sCsvFilePath
Dim lProcIdx
Dim lProcNum
lProcIdx = 0
lProcNum = cCsvFileList.Count
Call objPrgrsBar.Update(lProcIdx, lProcNum)
for each sCsvFilePath In cCsvFileList
    
    '*** 試験ログCSVオープン ***
    dim cFileContents
    Set cFileContents = CreateObject("System.Collections.ArrayList")
    call ReadTxtFileToCollection(sCsvFilePath, cFileContents)
    
    '*** 試験ログファイルチェック ***
    If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD And _
       Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
        
        '*** DataType取得 ***
        Dim vRamNames
        Dim vDataTypes
        vRamNames = Split(cFileContents(0), ",")
        vDataTypes = Split(cFileContents(1), ",")
        Dim sRamName
        Dim lIdx
        lIdx = 0
        for each sRamName In vRamNames
            If lIdx = 0 Then '1列目は無視
                'Do Nothing
            else
                if REPLACE_RAM_SYMBOL = True then
                    sRamName = RenameRamSymbol(sRamName)
                end if
                Dim sDataTypeListLine
                sDataTypeListLine = sRamName & "," & vDataTypes(lIdx)
                If Not dDataTypeListDupChk.Exists( sDataTypeListLine ) Then
                    cDataTypeList.Add sDataTypeListLine
                    dDataTypeListDupChk.Add sDataTypeListLine, ""
                end if
            end if
            lIdx = lIdx + 1
        next
    Else
        'Do Nothing
    End If
    
    lProcIdx = lProcIdx + 1
    Call objPrgrsBar.Update(lProcIdx, lProcNum)
    
    Set cFileContents = Nothing
next

'*****************************
' DataType一覧出力
'*****************************
call WriteTxtFileFrCollection(sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME, cDataTypeList, True)

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set cDataTypeList = Nothing
set dDataTypeListDupChk = Nothing

IF OUTPUT_FINISH_MESSAGE = True Then
    MsgBox "DataType一覧 生成完了!"
End If

'===============================================================================
' 関数
'===============================================================================
Private Function RenameRamSymbol( _
    byval sTrgtWord _
)
    Dim sOutWord
    sOutWord = sTrgtWord
    sOutWord = Replace(sOutWord, "[", "_")
    sOutWord = Replace(sOutWord, "]", "")
    RenameRamSymbol = sOutWord
End Function

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

