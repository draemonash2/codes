Attribute VB_Name = "Mng_Info_TCDoc_4UT"
Option Explicit

Private Const TC_SHT_KEY_WORD = "UT Checklist"
Private Const SRCH_KEYWORD_FILE_NAME = "File Name"
Private Const SRCH_KEYWORD_MODULE_NAME = "Module Name"
Private Const SRCH_KEYWORD_TESTER_SUMMARY = "UT実施者名"
Private Const SRCH_KEYWORD_TC_TESTER = "　評価者　"
Private Const SRCH_KEYWORD_TEST_RSLT = "結果判定"
Private Const SRCH_KEYWORD_REV = "Rev"
Private Const RSLT_SHEET_NAME = "Result(UT)"

Private Type T_TC_ROWCLM_INFO
    lTcNoClm As Long
    lTcDataClm As Long
    lTesterClm As Long
    lTestResultClm As Long
    lTestDateClm As Long
    lTestRevClm As Long
    lTcStrtRow As Long
    lTcEndRow As Long
End Type

Private Enum E_RSLT_SHT_ROW
    ROW_TC_DOC_NAME = 6
    ROW_LOG_DIR_PATH
    
    ROW_TITLE = 10
    ROW_TC_STRT
End Enum

Private Enum E_RSLT_SHT_CLM
    CLM_TC_DOC_NAME = 4
    CLM_LOG_DIR_PATH = 4
    
    CLM_PRE_STRT = 1  '開始列 - 1
    CLM_SHT_NAME
    CLM_TC_NO
    CLM_WRI_RSLT
    CLM_RESERVE_01 '未使用
    CLM_PRE_TESTER
    CLM_PRE_TEST_DATE
    CLM_PRE_TEST_RSLT
    CLM_PRE_TEST_DATA
    CLM_PRE_REV
    CLM_RESERVE_02 '未使用
    CLM_PST_TESTER
    CLM_PST_TEST_DATE
    CLM_PST_TEST_RSLT
    CLM_PST_TEST_DATA
    CLM_PST_REV
    CLM_PST_END '最終列 + 1
End Enum

Public Function GetTCDocInfo4UT( _
    ByRef wTrgtBook As Workbook _
)
    Dim lShtIdx As Long
    Dim lTcShtInfoIdx As Long
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Dim sSrchKeyWord As String
    Dim lRowIdx As Long
    Dim lTcNum As Long
    Dim tTcRowClmInfo As T_TC_ROWCLM_INFO
    
    With wTrgtBook
        '### 試験フェーズ取得 ###
        gtTestDocInfo.eTrgtPhase = TRGT_PHASE_UT
        
        '### ブック名取得 ###
        gtTestDocInfo.sTcDocName = .Name
        
        For lShtIdx = 1 To .Sheets.Count
            '項目シート判定
            tNearCellData = GetNearCellData(.Sheets(lShtIdx), TC_SHT_KEY_WORD, 0, 0)
            If tNearCellData.bIsCellDataExist = True Then
                If Sgn(gtTestDocInfo.atTcShtInfo) = 0 Then
                    lTcShtInfoIdx = 0
                Else
                    lTcShtInfoIdx = UBound(gtTestDocInfo.atTcShtInfo) + 1
                End If
                ReDim Preserve gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx)
                
                '==================
                '=== セル値取得 ===
                '==================
                '### シート名取得 ###
                gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName = .Sheets(lShtIdx).Name
                
                '### ソースファイル名取得 ###
                sSrchKeyWord = SRCH_KEYWORD_FILE_NAME
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sSrcFileName = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "以下のセルが見つかりませんでした！" & vbNewLine & _
                                        "  ブック名：" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  セル：" & sSrchKeyWord _
                                      )
                End If
                
                '### モジュール名取得 ###
                sSrchKeyWord = SRCH_KEYWORD_MODULE_NAME
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "以下のセルが見つかりませんでした！" & vbNewLine & _
                                        "  ブック名：" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  セル：" & sSrchKeyWord _
                                      )
                End If
                
                '### UT実施者名取得 ###
                sSrchKeyWord = SRCH_KEYWORD_TESTER_SUMMARY
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sTester = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "以下のセルが見つかりませんでした！" & vbNewLine & _
                                        "  ブック名：" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  セル：" & sSrchKeyWord _
                                      )
                End If
                
                '======================
                '=== セル列情報取得 ===
                '======================
                tTcRowClmInfo = GetCellClmInfo(wTrgtBook, lShtIdx)
                
                'エラー出力
                Call OutpErrorMsg(ERROR_PROC_STOP)
                
                '### 項目情報取得 ###
                '項目数０
                lTcNum = tTcRowClmInfo.lTcEndRow - tTcRowClmInfo.lTcStrtRow + 1
                If lTcNum = 0 Then
                    'Do Nothing
                Else
                    For lRowIdx = tTcRowClmInfo.lTcStrtRow To tTcRowClmInfo.lTcEndRow
                        Call GetTCDocInfoTestCase( _
                                                    wTrgtBook, _
                                                    lShtIdx, _
                                                    lTcShtInfoIdx, _
                                                    tTcRowClmInfo, _
                                                    lRowIdx _
                                                )
                    Next lRowIdx
                End If
            Else
                'Do Nothing
            End If
        Next lShtIdx
    End With
    
    Call OutpErrorMsg(ERROR_PROC_STOP)
End Function

Private Function GetCellClmInfo( _
    ByRef wTrgtBook As Workbook, _
    ByVal lShtIdx As Long _
) As T_TC_ROWCLM_INFO
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Dim sSrchKeyWord As String
    Dim tTcRowClmInfo As T_TC_ROWCLM_INFO
    
    With wTrgtBook
        '### 項番 セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_TC_NO
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTcNoClm = tNearCellData.lClm
        tTcRowClmInfo.lTcStrtRow = tNearCellData.lRow + 2
        tTcRowClmInfo.lTcEndRow = .Sheets(lShtIdx).Cells(.Sheets(lShtIdx).Rows.Count, tTcRowClmInfo.lTcNoClm).End(xlUp).Row
        
        '### 評価者 セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_TC_TESTER
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTesterClm = tNearCellData.lClm
        
        '### 年月日 セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_TEST_DATE
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTestDateClm = tNearCellData.lClm
        
        '### 結果判定 セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_TEST_RSLT
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTestResultClm = tNearCellData.lClm
        
        '### 試験データ セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_TEST_DATA
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTcDataClm = tNearCellData.lClm
        
        '### Rev セル情報取得 ###
        sSrchKeyWord = SRCH_KEYWORD_REV
        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
        If tNearCellData.bIsCellDataExist = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                "  ブック名：" & .Name & vbNewLine & _
                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                "  セル：" & sSrchKeyWord _
                              )
        End If
        tTcRowClmInfo.lTestRevClm = tNearCellData.lClm
    End With
    GetCellClmInfo = tTcRowClmInfo
End Function

Private Function GetTCDocInfoTestCase( _
    ByRef wTrgtBook As Workbook, _
    ByVal lShtIdx As Long, _
    ByVal lTcShtInfoIdx As Long, _
    ByRef tTcRowClmInfo As T_TC_ROWCLM_INFO, _
    ByVal lRowIdx As Long _
)
    Dim sTcNo As String
    Dim sTester As String
    Dim sTestDate As String
    Dim sTestResult As String
    Dim sRev As String
    Dim sTestDataCellValue As String
    Dim asTestDataFileNames() As String
    Dim sTestDataFileName As String
    Dim lTestDataIdx As Long
    Dim sDicKey As String
    Dim sDicItem As String
    Dim lTcInfoIdx As Long
    
    With wTrgtBook
        sTcNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm).Value
        '項番が「-」か空欄の場合、無視
        If sTcNo = "-" Or sTcNo = "" Then
            'Do Nothing
        Else
            sTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
            sTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
            sTestResult = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
            sRev = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
            sTestDataCellValue = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
            
            If Sgn(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo) = 0 Then
                lTcInfoIdx = 0
            Else
                lTcInfoIdx = UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo) + 1
            End If
            ReDim Preserve gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx)
            
            '### 項番取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo = sTcNo
            
            '### 評価者取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTester = sTester
            
            '### 年月日取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDate = sTestDate
            
            '### 結果判定取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestResult = sTestResult
            
            '### Rev取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevSrc = sRev
            
            '### 試験データ＋期待ファイルパスリスト取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue = sTestDataCellValue
            If sTestDataCellValue = "" Then
                'Do Nothing
            Else
                'セル内の改行をデリミタとして配列に分解
                asTestDataFileNames = GetTestDataArray(sTestDataCellValue)
                
                ReDim Preserve gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(UBound(asTestDataFileNames))
                For lTestDataIdx = 0 To UBound(asTestDataFileNames)
                    sTestDataFileName = asTestDataFileNames(lTestDataIdx)
                    '試験データ
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestDataIdx) = sTestDataFileName
                    
                    '期待ＸＸＸファイルパスリスト
                    If sTestDataFileName <> "-" Then
                        Select Case GetFileNameExt(sTestDataFileName)
                            Case "csv"
                                sDicKey = GetFileNameBase(.Name) & "\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                          sTestDataFileName
                                sDicItem = lTcShtInfoIdx & "_" & lTcInfoIdx & "_" & lTestDataIdx
                                If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                                    'Debug.Print "重複Key：" & sDicKey
                                Else
                                    gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                                End If
                            Case "htm"
                                sDicKey = GetFileNameBase(.Name) & "\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                          sTestDataFileName
                                sDicItem = lTcShtInfoIdx
                                If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                                    'Debug.Print "重複Key：" & sDicKey
                                Else
                                    gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                                End If
                            Case "txt"
                                sDicKey = GetFileNameBase(.Name) & "\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                          "TestCoverLog\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sSrcFileName & "\" & _
                                          sTestDataFileName
                                sDicItem = lTcShtInfoIdx
                                If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                                    'Debug.Print "重複Key：" & sDicKey
                                Else
                                    gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                                End If
                            Case Else
                                'ここではチェックしない。
                        End Select
                    Else
                        'Do Nothing
                    End If
                Next lTestDataIdx
            End If
        End If
    End With
End Function

Public Function OutpTestResult4UT( _
    ByRef wTcDocBook As Workbook, _
    ByRef wWriRsltBook As Workbook _
)
    Dim oTestedTcList As Object 'Key:[シート名]_[項番]、Item:未使用（True固定）
    
    '=== 試験実施済み項目リスト作成 ===
    Set oTestedTcList = CreTestedTcList
    
    '=== 試験結果欄書き込み ===
    Call WriTestRslt(wTcDocBook, oTestedTcList)
    
    '=== 試験結果欄書き込み結果出力 ===
    Call OutpWriRslt(wWriRsltBook)
End Function

Private Function CreTestedTcList() As Object
    Dim sPath As String
    Dim sRelativeFilePath As String
    Dim eSvnModStatus As E_SVN_MOD_STATUS
    Dim sFileBaseName As String
    Dim atSvnModStatInfo() As T_SVN_MOD_STAT_INFO
    Dim lSvnModStatInfoIdx As Long
    Dim oTestedTcList As Object 'Key:[シート名]_[項番]、Item:未使用（True固定）
    
    Set oTestedTcList = CreateObject("Scripting.Dictionary")
    
    'SVN の変更状態リスト取得
    atSvnModStatInfo = GetSvnModStatList(gtInputInfo.sTestLogDirPath)
    
    For lSvnModStatInfoIdx = 0 To UBound(atSvnModStatInfo)
    '変更済み状態リスト
        sPath = atSvnModStatInfo(lSvnModStatInfoIdx).sPath
        eSvnModStatus = atSvnModStatInfo(lSvnModStatInfoIdx).eSvnModStat
        sRelativeFilePath = Replace(sPath, gtInputInfo.sTestLogDirPath & "\", "")
        
        If GetTypeFileOrFolder(sPath) <> PATH_TYPE_FILE Then
            'Do Nothing
        Else
            If GetFileNameExt(sRelativeFilePath) <> "csv" Then
                'Do Nothing
            Else
                If eSvnModStatus = MOD_STAT_NOTCHANGE Then
                    'Do Nothing
                Else
                    '試験実施済みリスト追加
                    sFileBaseName = GetFileNameBase(sRelativeFilePath)
                    Debug.Assert InStr(sFileBaseName, "_") > 0
                    oTestedTcList.Add _
                        Split(sFileBaseName, "_")(0) & "_" & Split(sFileBaseName, "_")(1), _
                        True
                End If
            End If
        End If
    Next
    
    Set CreTestedTcList = oTestedTcList
End Function

Private Function WriTestRslt( _
    ByRef wTcDocBook As Workbook, _
    ByRef oTestedTcList As Object _
)
    Dim lShtIdx As Long
    Dim lTestCaseIdx As Long
    Dim tTcRowClmInfo As T_TC_ROWCLM_INFO
    Dim lTcNum As Long
    Dim sTestCaseNo As String
    Dim sSheetName As String
    Dim sModuleName As String
    Dim sTester As String
    Dim sTestDate As String
    Dim sTestRslt As String
    Dim sTcData As String
    Dim sRevSrc As String
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Dim sSrchKeyWord As String
    Dim lRowIdx As Long
    Dim lWriRsltInfoRowIdx As Long
    Dim bIsTested As Boolean
    
    '*** 書き込み結果（試験項目書名、ログフォルダパス）格納 ***
    gtWriRsltInfo.sTcDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    gtWriRsltInfo.sLogDirPath = gtInputInfo.sTestLogDirPath
    
    With wTcDocBook
        For lShtIdx = 1 To .Sheets.Count
            '項目シート判定
            tNearCellData = GetNearCellData(.Sheets(lShtIdx), TC_SHT_KEY_WORD, 0, 0)
            If tNearCellData.bIsCellDataExist = True Then
                '各項目の列番号取得
                tTcRowClmInfo = GetCellClmInfo(wTcDocBook, lShtIdx)
                
                '### 項目情報取得 ###
                '項目数０
                lTcNum = tTcRowClmInfo.lTcEndRow - tTcRowClmInfo.lTcStrtRow + 1
                If lTcNum = 0 Then
                    'Do Nothing
                Else
                    sSheetName = .Sheets(lShtIdx).Name
                    
                    'モジュール名
                    sSrchKeyWord = SRCH_KEYWORD_MODULE_NAME
                    tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                    If tNearCellData.bIsCellDataExist = True Then
                        sModuleName = tNearCellData.sCellValue
                    Else
                        Call StoreErrorMsg( _
                                            "以下のセルが見つかりませんでした！" & vbNewLine & _
                                            "  ブック名：" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                            "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                            "  セル：" & sSrchKeyWord _
                                          )
                    End If
                    
                    bIsTested = False
                    For lRowIdx = tTcRowClmInfo.lTcStrtRow To tTcRowClmInfo.lTcEndRow
                        If Sgn(gtWriRsltInfo.atWriRsltInfoRow) = 0 Then
                            lWriRsltInfoRowIdx = 0
                        Else
                            lWriRsltInfoRowIdx = UBound(gtWriRsltInfo.atWriRsltInfoRow) + 1
                        End If
                        ReDim Preserve gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx)
                        
                        sTestCaseNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm)
                        '試験済みか判定
                        If oTestedTcList.exists(sSheetName & "_" & sTestCaseNo) = True Then
                            bIsTested = True
                            sTester = gtInputInfo.sTester
                            sTestDate = gtInputInfo.sTestDate
                            sTestRslt = gtInputInfo.sTestRslt
                            sTcData = _
                                sSheetName & "_" & sTestCaseNo & ".csv" & vbLf & _
                                sModuleName & ".txt" & vbLf & _
                                "テスト結果報告書.htm"
                            sRevSrc = gtInputInfo.sRevSrc
                            
                            '*** 書き込み結果格納 ***
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sSheetName = sSheetName
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sTestCaseNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm).Value
                            If .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value = "" Then
                                gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "新規"
                            Else
                                gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "上書き"
                            End If
                            '書き込み前
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            '書き込み後
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTester = sTester
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestDate = sTestDate
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt = sTestRslt
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestData = sTcData
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostRevSrc = sRevSrc
                            
                            '*** 試験データ欄書き込み ***
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value = sTester
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value = sTestDate
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value = sTestRslt
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value = sTcData
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value = sRevSrc
                        Else
                            '*** 書き込み結果格納 ***
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sSheetName = sSheetName
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sTestCaseNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "変更なし"
                            '書き込み前
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            '書き込み後
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            
                            '*** 試験データ欄書き込み ***
                            '試験済みでないため、書き換えない
                        End If
                    Next lRowIdx
                    
                    '一項目でも試験済みであれば
                    If bIsTested = True Then
                        'UT実施者名
                        sSrchKeyWord = SRCH_KEYWORD_TESTER_SUMMARY
                        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                        If tNearCellData.bIsCellDataExist = True Then
                            .Sheets(lShtIdx).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = gtInputInfo.sTester
                        Else
                            Call StoreErrorMsg( _
                                                "以下のセルが見つかりませんでした！" & vbNewLine & _
                                                "  ブック名：" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                                "  シート名：" & .Sheets(lShtIdx).Name & vbNewLine & _
                                                "  セル：" & sSrchKeyWord _
                                              )
                        End If
                    Else
                        'Do Nothing
                    End If
                End If
            Else
                'Do Nothing
            End If
        Next lShtIdx
    End With
End Function

'★★★未テスト！★★★
Private Function OutpWriRslt( _
    ByRef wWriRsltBook As Workbook _
)
    Dim lWriRsltInfoRowIdx As Long
    Dim lRowIdx As Long
    
    '+++ シートコピー +++
    Call CopyRsltSht(wWriRsltBook, RSLT_SHEET_NAME)
    
    '+++ 書き込み結果出力 +++
    With wWriRsltBook.Sheets(RSLT_SHEET_NAME)
        '### 試験項目書ファイル名 ###
        .Cells(ROW_TC_DOC_NAME, CLM_TC_DOC_NAME).Value = gtWriRsltInfo.sTcDocFileName
        
        '### 試験ログフォルダ名 ###
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtWriRsltInfo.sLogDirPath
        
        '### 行情報 ###
        If Sgn(gtWriRsltInfo.atWriRsltInfoRow) = 0 Then
            'Do Nothing
        Else
            For lWriRsltInfoRowIdx = 0 To UBound(gtWriRsltInfo.atWriRsltInfoRow)
                lRowIdx = ROW_TC_STRT + lWriRsltInfoRowIdx
                
                .Cells(lRowIdx, CLM_SHT_NAME).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sSheetName
                .Cells(lRowIdx, CLM_TC_NO).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sTestCaseNo
                .Cells(lRowIdx, CLM_WRI_RSLT).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt
                
                .Cells(lRowIdx, CLM_PRE_TESTER).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTester
                .Cells(lRowIdx, CLM_PRE_TEST_DATE).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestDate
                .Cells(lRowIdx, CLM_PRE_TEST_RSLT).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt
                .Cells(lRowIdx, CLM_PRE_TEST_DATA).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestData
                .Cells(lRowIdx, CLM_PRE_REV).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreRevSrc
                
                .Cells(lRowIdx, CLM_PST_TESTER).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTester
                .Cells(lRowIdx, CLM_PST_TEST_DATE).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestDate
                .Cells(lRowIdx, CLM_PST_TEST_RSLT).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt
                .Cells(lRowIdx, CLM_PST_TEST_DATA).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestData
                .Cells(lRowIdx, CLM_PST_REV).Value = gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostRevSrc
            Next lWriRsltInfoRowIdx
        End If
    End With
    
End Function

'★★★テスト用★★★
Sub test()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\test\11_TF2次プロト2nd\31_単体試験ログ\02_診断モニタ変更"
    Call OutpTestResult4UT(Nothing, Nothing)
End Sub
