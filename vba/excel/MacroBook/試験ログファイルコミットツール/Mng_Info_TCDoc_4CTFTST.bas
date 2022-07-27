Attribute VB_Name = "Mng_Info_TCDoc_4CTFTST"
Option Explicit

Private Const TC_SHT_KEY_WORD = "試験項目"
Private Const SRCH_KEYWORD_TC_TESTER = "評価者"
Private Const SRCH_KEYWORD_TEST_RSLT = "判定"
Private Const SRCH_KEYWORD_REV_HEXABS = "HEX/ABS"
Private Const SRCH_KEYWORD_REV_A2L = "A2L"

Private Type T_TC_CLM_INFO
    lTcNoClm As Long
    lTcDataClm As Long
    lTesterClm As Long
    lTestResultClm As Long
    lTestDateClm As Long
    lTestRevHexAbsClm As Long
    lTestRevA2LClm As Long
End Type

Public Function GetTCDocInfo4CTFTST( _
    ByRef wTrgtBook As Workbook _
)
    Dim lShtIdx As Long
    Dim lTcShtInfoIdx As Long
    Dim lTcStrtRow As Long
    Dim lTcEndRow As Long
    Dim tTcClmInfo As T_TC_CLM_INFO
    Dim lRowIdx As Long
    Dim lTcNum As Long
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Dim sSrchKeyWord As String
    
    With wTrgtBook
        '### 試験フェーズ取得 ###
        gtTestDocInfo.eTrgtPhase = gtInputInfo.eTrgtPhase
        
        '### ブック名取得 ###
        gtTestDocInfo.sTcDocName = wTrgtBook.Name
        
        For lShtIdx = 1 To wTrgtBook.Sheets.Count
            '項目シート判定
            If InStr(wTrgtBook.Sheets(lShtIdx).Name, TC_SHT_KEY_WORD) > 0 Then
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
                gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName = wTrgtBook.Sheets(lShtIdx).Name
                
                '====================
                '=== セル情報取得 ===
                '====================
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
                tTcClmInfo.lTcNoClm = tNearCellData.lClm
                lTcStrtRow = tNearCellData.lRow + 2
                lTcEndRow = .Sheets(lShtIdx).Cells(.Sheets(lShtIdx).Rows.Count, tTcClmInfo.lTcNoClm).End(xlUp).Row
                lTcNum = lTcEndRow - lTcStrtRow + 1
                
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
                tTcClmInfo.lTesterClm = tNearCellData.lClm
                
                '### 判定 セル情報取得 ###
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
                tTcClmInfo.lTestResultClm = tNearCellData.lClm
                
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
                tTcClmInfo.lTestDateClm = tNearCellData.lClm
                
                '### Rev（HEX/ABS） セル情報取得 ###
                sSrchKeyWord = SRCH_KEYWORD_REV_HEXABS
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
                tTcClmInfo.lTestRevHexAbsClm = tNearCellData.lClm
                
                '### Rev（A2l） セル情報取得 ###
                sSrchKeyWord = SRCH_KEYWORD_REV_A2L
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
                tTcClmInfo.lTestRevA2LClm = tNearCellData.lClm
                
                '### 試験データセル情報取得 ###
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
                tTcClmInfo.lTcDataClm = tNearCellData.lClm
                
                'エラー出力
                Call OutpErrorMsg(ERROR_PROC_STOP)
                
                '### 項目情報取得 ###
                '項目数０
                If lTcNum = 0 Then
                    'Do Nothing
                Else
                    For lRowIdx = lTcStrtRow To lTcEndRow
                        Call GetTCDocInfoTestCase( _
                                                    wTrgtBook, _
                                                    lShtIdx, _
                                                    lTcShtInfoIdx, _
                                                    tTcClmInfo, _
                                                    lRowIdx _
                                                )
                    Next lRowIdx
                End If
                
                '単体試験用の情報に何も格納されていないことを確認
                Debug.Assert gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName = ""
                Debug.Assert gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sSrcFileName = ""
            Else
                'Do Nothing
            End If
        Next lShtIdx
    End With
    
    Call OutpErrorMsg(ERROR_PROC_STOP)
End Function

Private Function GetTCDocInfoTestCase( _
    ByRef wTrgtBook As Workbook, _
    ByVal lShtIdx As Long, _
    ByVal lTcShtInfoIdx As Long, _
    ByRef tTcClmInfo As T_TC_CLM_INFO, _
    ByVal lRowIdx As Long _
)
    Dim sTcNo As String
    Dim sTester As String
    Dim sTestDate As String
    Dim sTestResult As String
    Dim sTestRevHexAbs As String
    Dim sTestRevA2L As String
    Dim sTcData As String
    Dim sTestDataCellValue As String
    Dim asTestDataFileNames() As String
    Dim sTestDataFileName As String
    Dim lTestDataIdx As Long
    Dim sDicKey As String
    Dim sDicItem As String
    Dim lTcInfoIdx As Long
    
    With wTrgtBook
        sTcNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTcNoClm).Value
        '項番が「-」か空欄の場合、無視
        If sTcNo = "-" Or sTcNo = "" Then
            'Do Nothing
        Else
            sTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTesterClm).Value
            sTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTestDateClm).Value
            sTestResult = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTestResultClm).Value
            sTestRevHexAbs = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTestRevHexAbsClm).Value
            sTestRevA2L = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTestRevA2LClm).Value
            sTestDataCellValue = .Sheets(lShtIdx).Cells(lRowIdx, tTcClmInfo.lTcDataClm).Value
            
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
            
            '### 判定取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestResult = sTestResult
            
            '### Rev（HEX/ABS）取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevHexAbs = sTestRevHexAbs
            
            '### Rev（A2L）取得 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevA2L = sTestRevA2L
            
            '### 試験データ＋期待ＸＸＸファイルパスリスト取得 ###
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
                        sDicKey = GetFileNameBase(.Name) & "\" & _
                                  gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                  sTestDataFileName
                        sDicItem = lTcShtInfoIdx & "_" & lTcInfoIdx & "_" & lTestDataIdx
                        If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                            'Debug.Print "重複Key：" & sDicKey
                        Else
                            gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                        End If
                    Else
                        'Do Nothing
                    End If
                Next lTestDataIdx
            End If
        End If
    End With
End Function

