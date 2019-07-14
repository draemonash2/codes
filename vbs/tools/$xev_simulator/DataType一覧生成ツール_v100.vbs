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
'         → REPLACE_RAM_NAME = True:有効 / False:無効
'     - 整形完了時のメッセージ出力有無
'         → OUTPUT_FINISH_MESSAGE = True:出力 / False:出力しない
'
'【改訂履歴】
'   1.0.0   2019/07/01  遠藤    ・新規作成
'===============================================================================
'===============================================================================
' 設定
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST REPLACE_RAM_NAME = False
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
                if REPLACE_RAM_NAME = True then
                    sRamName = ReplaceKeyword(sRamName)
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
Private Function ReplaceKeyword( _
    byval sTrgtWord _
)
    Dim sOutWord
    sOutWord = sTrgtWord
    sOutWord = Replace(sOutWord, "[", "_")
    sOutWord = Replace(sOutWord, "]", "")
    ReplaceKeyword = sOutWord
End Function

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する
' = 引数    sTrgtDir        String      [in]    対象フォルダ
' = 引数    cFileList       Collections [out]   ファイル/フォルダパス一覧
' = 引数    lFileListType   Long        [in]    取得する一覧の形式
' =                                                 0：両方
' =                                                 1:ファイル
' =                                                 2:フォルダ
' =                                                 それ以外：格納しない
' = 戻値    なし
' = 覚書    ・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' =         ・Arrayコレクションに格納する
' ==================================================================
Public Function GetFileList3( _
    ByVal sTrgtDir, _
    ByRef cFileList, _
    ByVal lFileListType _
)
    Dim objFSO  'FileSystemObjectの格納先
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
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        WScript.Echo "エラー " & Err.Description
    End If  
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error Goto 0
End Function
'   Call Test_GetFileList3()
    Private Sub Test_GetFileList3()
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim sCurDir
        sCurDir = objFSO.GetParentFolderName( WScript.ScriptFullName )
        
        msgbox sCurDir
        
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        Call GetFileList3( sCurDir, cFileList, 1 )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox sOutput
    End Sub

' ==================================================================
' = 概要    指定ファイルパスが存在する場合、"_XXX" を付与して返却する
' = 引数    sTrgtFilePath   String      [in]    対象パス
' = 戻値                    String              付与後パス
' = 覚書    本関数では、ファイルは作成しない。
' = 依存lib なし
' ==================================================================
Public Function GetFileNotExistPath( _
    ByVal sTrgtFilePath _
)
    Dim lIdx
    Dim objFSO
    Dim sFileParDirPath
    Dim sFileBaseName
    Dim sFileExtName
    Dim sCreFilePath
    Dim bIsTrgtPathExists
    
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreFilePath = sTrgtFilePath
    bIsTrgtPathExists = False
    Do While objFSO.FileExists( sCreFilePath )
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sFileParDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
        sFileBaseName = objFSO.GetBaseName( sTrgtFilePath ) & "_" & String( 3 - len(lIdx), "0" ) & lIdx
        sFileExtName = objFSO.GetExtensionName( sTrgtFilePath )
        If sFileExtName = "" Then
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName
        Else
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName & "." & sFileExtName
        End If
    Loop
    GetFileNotExistPath = sCreFilePath
End Function
'   Call Test_GetFileNotExistPath()
    Private Sub Test_GetFileNotExistPath()
        Dim sOutStr
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

' ==================================================================
' = 概要    テキストファイルの中身を配列に格納
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [out]   ファイルの中身
' = 戻値    読み出し結果    Boolean             読み出し結果
' =                                                 True:ファイル存在
' =                                                 False:それ以外
' = 覚書    なし
' ==================================================================
Public Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.add objTxtFile.ReadLine
            Loop
            ReadTxtFileToCollection = True
        Else
            ReadTxtFileToCollection = False
        '   WScript.Echo "エラー " & Err.Description
        End If
        
        objTxtFile.Close
    Else
        ReadTxtFileToCollection = False
    End If
    On Error Goto 0
End Function
'   Call Test_OpenTxtFile2Array()
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        sFilePath = "C:\codes\vbs\試験結果CSV整形ツール\data_type_list_.csv"
        Dim bRet
        bRet = ReadTxtFileToCollection( sFilePath, cFileList )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox bRet
        MsgBox sOutput
    End Sub

' ==================================================================
' = 概要    配列の中身をテキストファイルに書き出し
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [in]    ファイルの中身
' = 引数    bOverwrite      Boolean     [in]    True:上書き、False:新規ファイル
' = 戻値    書き出し結果    Boolean             書き出し結果
' =                                                 True:書き出し成功
' =                                                 False:それ以外
' = 覚書    なし
' = 依存lib FileSystem.vbs/GetFileNotExistPath
' ==================================================================
Public Function WriteTxtFileFrCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents, _
    ByVal bOverwrite _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objTxtFile
    If bOverwrite = True Then
        'Do Nothing
    Else
        Dim sInTrgtFilePath
        sInTrgtFilePath = sTrgtFilePath
        sTrgtFilePath = GetFileNotExistPath(sInTrgtFilePath)
    End If
    Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
    
    If Err.Number = 0 Then
        Dim sFileLine
        For Each sFileLine In cFileContents
            objTxtFile.WriteLine sFileLine
        Next
        WriteTxtFileFrCollection = True
    Else
        WriteTxtFileFrCollection = False
    '   WScript.Echo "エラー " & Err.Description
    End If
    
    objTxtFile.Close
    On Error Goto 0
End Function
'   Call Test_WriteTxtFileFrCollection()
    Private Sub Test_WriteTxtFileFrCollection()
        Dim cFileContents
        Set cFileContents = CreateObject("System.Collections.ArrayList")
        cFileContents.Add "a"
        cFileContents.Add "b"
        cFileContents.Add "d"
        cFileContents.Add "e"
        cFileContents.Insert 1, "c"
        DIm sTrgtFilePath
        sTrgtFilePath = "C:\codes\vbs\_lib\Test.csv.bak"
        call WriteTxtFileFrCollection( sTrgtFilePath, cFileContents, False )
    End Sub

' progrress bar cscript class v1.00
Class ProgressBar
    Private sStatus
    Private objFSO
    Private objWshShell
    
    Private Sub Class_Initialize
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        Dim sExeFileName
        sExeFileName = LCase(objFSO.GetFileName(WScript.FullName))
        if sExeFileName = "cscript.exe" then
            'Do Nothing
        else
            objWshShell.Run "cscript //nologo """ & Wscript.ScriptFullName & """", 1, False
            Wscript.Quit
        end if
    End Sub
    
    Private Sub Class_Terminate
        Set objFSO = Nothing
        Set objWshShell = Nothing
    End Sub
    
    ' ==================================================================
    ' = 概要    メッセージを更新する
    ' = 引数    sProgMsg      String   [in] メッセージ
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Message( _
        ByVal sMessage _
    )
        if sStatus = "Update" then
            Wscript.StdOut.Write vbCrLf
        end if
        Wscript.StdOut.Write sMessage & vbCrLf
        sStatus = "Message"
    End Property
    
    ' ==================================================================
    ' = 概要    進捗を更新する
    ' = 引数    lBunsi      Long   [in] 進捗
    ' = 引数    lBunbo      Long   [in] 進捗最大値
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Sub Update( _
        ByVal lBunsi, _
        ByVal lBunbo _
    )
        'パーセンテージ計算
        Dim iPercentage
        Dim sPercentage
        iPercentage = Cint((lBunsi / lBunbo) * 100)
        sPercentage = iPercentage & "%"
        sPercentage = String(4 - Len(sPercentage), " ") & sPercentage
        
        '進捗バー
        Dim sProgressBar
        sProgressBar = String(Cint(iPercentage/5), "=") & ">" & String(20 - Cint(iPercentage/5), " ")
        
        '描画
        Wscript.StdOut.Write sPercentage & " |" & sProgressBar & "| " & lBunsi & "/" & lBunbo & vbCr
        sStatus = "Update"
    End Sub
    
    ' ==================================================================
    ' = 概要    プログレスバーを終了する
    ' = 引数    なし
    ' = 戻値    なし
    ' = 覚書    cscriptは終了できない
    ' ==================================================================
'   Public Function Quit()
'       gobjExplorer.Document.Body.Style.Cursor = "default"
'       gobjExplorer.Quit
'   End Function
    
End Class
    If WScript.ScriptName = "ProgressBarCscript.vbs" Then
        Call Test_ProgressBar
    End If
    Private Sub Test_ProgressBar
        Dim lProcIdx
        Dim lProcNum
        Dim objPrgrsBar
        Set objPrgrsBar = New ProgressBar
        
        '#処理１
        objPrgrsBar.Message = "長い処理 実行!"
        lProcNum = 255
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 1
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        '#処理２
        objPrgrsBar.Message = "短い処理 実行!"
        lProcNum= 10
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 45
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        objPrgrsBar.Message = "Complete!!"
        msgbox "終了しました"
    End Sub
