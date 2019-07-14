Option Explicit

'===============================================================================
'【概要】
'   xEVシミュレータが出力した試験ログCSVを整形し、CANapeでインポートできる形式に変換する。
'       ・「Datatype」列を付与
'           （DataTypeは data_type_list.csv より取得）
'       ・RAM名から配列識別子を除去
'           ex) ram[0]:1 → ram_0:1
'
'【使用方法】
'   使用方法は２通り。
'       ◆フォルダ配下の全csvすべてを置換したい場合
'           １．「data_type_list.csv」を作成。
'                 ex) AAA:1[1],uint8
'                     AAA:1[2],uint8
'                     BBB:1,sint16
'                     CCC:2,double
'           ２．整形対象の試験ログ(CSV)と同じフォルダに
'               「試験ログCSV整形ツール.vbs」と「data_type_list.csv」を格納。
'           ３．「試験ログCSV整形ツール.vbs」を実行。
'               （ダブルクリック or コマンドプロンプトで実行）
'       ◆１ファイルのみ整形したい場合
'           １．「data_type_list.csv」を作成。
'           ２. 整形したい試験ログ(CSV)を「試験ログCSV整形ツール.vbs」へdrag&dropする。
'
'【詳細仕様】
'   ・ファイルの先頭に"TimeStamp"と記載された.csvファイルを試験ログ(CSV)と解釈する。
'   ・以下のような追加設定が可能。
'     - RAM名から配列識別子を除去する機能の有効無効
'         → REPLACE_RAM_NAME = True:有効 / False:無効
'     - 試験ログ(CSV)のバックアップを作成有無
'         → CREATE_BACKUP_FILE = True:バックアップファイル作成 / False:上書き
'     - 整形完了時のメッセージ出力有無
'         → FINISH_MESSAGE_OUTPUT = True:出力 / False:出力しない
'   ・data_type_list.csv について
'     - data_type_list.csv が存在しない場合は、すべて uint8 と解釈する。
'     - data_type_list.csv に存在しないRAMは、uint8 と解釈する。
'
'【改訂履歴】
'   1.0.0   2019/07/01  遠藤    ・新規作成
'===============================================================================

'===============================================================================
'= インクルード
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\String.vbs" )              'GetFileExt()
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )          'GetFileList3()
                                                            'GetFileNotExistPath()
Call Include( "C:\codes\vbs\_lib\Collection.vbs" )          'ReadTxtFileToCollection()
                                                            'WriteTxtFileFrCollection()
Call Include( "C:\codes\vbs\_lib\ProgressBarCscript.vbs" )  'Class ProgressBar

'===============================================================================
' 設定
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST DEFAULT_DATA_TYPE = "uint8"
CONST CREATE_BACKUP_FILE = False
CONST REPLACE_RAM_NAME = False
CONST FINISH_MESSAGE_OUTPUT = True

'===============================================================================
' 本処理
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp,"
Const DATATYPE_ROW_KEYWORD = "DataType"

Dim objPrgrsBar
Set objPrgrsBar = New ProgressBar
objPrgrsBar.Message = "試験ログCSV整形中..."

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
' DataType一覧取得
'*****************************
dim dDataTypeList
Set dDataTypeList = CreateObject("Scripting.Dictionary")

Dim sDataTypeListFilePath
sDataTypeListFilePath = sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME

dim objTxtFile
If objFSO.FileExists(sDataTypeListFilePath) Then
    set objTxtFile = objFSO.OpenTextFile(sDataTypeListFilePath, 1)

    dim objWords
    Dim sTxtLine
    Do Until objTxtFile.AtEndOfStream
        sTxtLine = objTxtFile.ReadLine
        objWords = split(sTxtLine, ",")
        if REPLACE_RAM_NAME = True then
            objWords(0) = ReplaceKeyword(objWords(0))
        end if
        On Error Resume Next '重複キーがあったら無視
        dDataTypeList.Add objWords(0), objWords(1) 'RamName DataType
        On Error Goto 0
    Loop
    objTxtFile.Close
Else
    'Do Nothing
End If

'*****************************
' 試験ログCSV整形
'*****************************
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
    If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD Then
        
        '*** バックアップ出力 ***
        If CREATE_BACKUP_FILE = True then
            Dim sCsvBakFilePathRaw
            Dim sCsvBakFilePath
            Dim lBakFileIdx
            sCsvBakFilePathRaw = sCsvFilePath & ".bak"
            sCsvBakFilePath = sCsvBakFilePathRaw
            lBakFileIdx = 1
            Do While objFSO.FileExists( sCsvBakFilePath )
                sCsvBakFilePath = sCsvBakFilePathRaw & lBakFileIdx
                lBakFileIdx = lBakFileIdx + 1
            Loop
            objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
        End If
        
        '*** 変数名置換 ***
        if REPLACE_RAM_NAME = True then
            cFileContents(0) = ReplaceKeyword(cFileContents(0))
        end if
        
        '*** Datatype置換or挿入 ***
        Dim vRamNames
        vRamNames = Split(cFileContents(0), ",")
        Dim sRamName
        Dim sDataTypeLine
        Dim lIdx
        lIdx = 0
        for each sRamName In vRamNames
            If lIdx = 0 Then '1列目は無視
                sDataTypeLine = DATATYPE_ROW_KEYWORD
            else
                'すでに置換済み
                'if REPLACE_RAM_NAME = True then
                '   sRamName = ReplaceKeyword(sRamName)
                'end if
                if dDataTypeList.Exists(sRamName) then
                    sDataTypeLine = sDataTypeLine & "," & dDataTypeList.Item(sRamName)
                else
                    sDataTypeLine = sDataTypeLine & "," & DEFAULT_DATA_TYPE
                end if
            end if
            lIdx = lIdx + 1
        next
        If Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
            cFileContents(1) = sDataTypeLine
        Else
            cFileContents.Insert 1, sDataTypeLine
        End If
        
        '*** CSV出力 ***
        call WriteTxtFileFrCollection(sCsvFilePath, cFileContents, True)
    Else
        'Do Nothing
    End If
    
    lProcIdx = lProcIdx + 1
    Call objPrgrsBar.Update(lProcIdx, lProcNum)
    
    Set cFileContents = Nothing
next

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set dDataTypeList = Nothing

IF FINISH_MESSAGE_OUTPUT = True Then
    MsgBox "試験ログCSV 整形完了!"
End If

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
