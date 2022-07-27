Attribute VB_Name = "Chk_A_TestData"
Option Explicit

Private Const SHEET_NAME = "(A)"

Private Enum E_ROW
    ROW_TC_DOC_NAME = 6
    ROW_RESERVE_01 '未使用
    ROW_ERR_NUM
    ROW_WARN_NUM
    ROW_RESERVE_02 '未使用
    ROW_ERR_MSG_TITLE
    ROW_ERR_MSG_STRT
End Enum

Private Enum E_CLM
    CLM_TC_DOC_NAME = 4
    CLM_TC_ERR_NUM = 4
    CLM_TC_WARN_NUM = 4
    
    CLM_STRT_PRE_CLM = 2 '最前列 - 1
    CLM_SHT_NAME
    CLM_TC_NO
    CLM_TC_DATA
    CLM_RESERVE_01 '未使用
    CLM_CHK_RSLT
    CLM_CHK_DETAIL
    CLM_END_NEXT_CLM '最後列 + 1
End Enum

Private Type T_ERROR_MSG_ROW
    sBookName As String  '現状未使用
    sSheetName As String
    sTcNo As String
    sTcData As String
    sChkRslt As String
    sErrDetail As String
End Type

Private Type T_ERROR_MSG
    sDocFileName As String
    lErrNum As Long
    lWarningNum As Long
    atErrorMsgRow() As T_ERROR_MSG_ROW
End Type

Private gtErrMsg As T_ERROR_MSG

Public Function ChkTestDataInit()
    Dim tErrMsg As T_ERROR_MSG
    gtErrMsg = tErrMsg
End Function

Public Function ChkTestDataMain( _
    ByRef wChkRsltBook As Workbook _
)
    Call StoreErrMsg
    Call OutpErrMsg(wChkRsltBook)
End Function

Private Function StoreErrMsg()
    '### タイトル部格納 ###
    gtErrMsg.sDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    'gtErrMsg.lErrNum 'エラー数はエラーメッセージ部格納処理にて格納する
    
    '### エラーメッセージ部格納 ###
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
        Call StoreErrMsgUT
    Else
        Call StoreErrMsgExceptUT
    End If
    
    '結果まとめシートへエラー数/ワーニング数を通知
    Call SetErrNum2Summary(FUNC_TYPE_A, gtErrMsg.lErrNum, gtErrMsg.lWarningNum)
End Function

Private Function StoreErrMsgUT()
    Dim lTcShtInfoIdx As Long
    Dim lTcInfoIdx As Long
    Dim lTestLogNameIdx As Long
    Dim bIsExistCsvFile As Boolean
    Dim bIsExistTxtFile As Boolean
    Dim bIsExistHtmFile As Boolean
    Dim sTestLogFileName As String
    Dim sTestLogBaseName As String
    Dim sTestLogExt As String
    Dim sErrDetail As String
    Dim sSheetName As String
    Dim sModuleName As String
    Dim sTestCaseNo As String
    Dim sTestDataCellValue As String
    Dim lErrMsgIdx As Long
    
    lErrMsgIdx = 0
    For lTcShtInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo)
        sModuleName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName
        sSheetName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName
        For lTcInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo)
            ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
            sTestCaseNo = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo
            
            '### シート名格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
            
            '### 項番格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
            
            '### 試験データ欄格納 ###
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
            
            If sTestDataCellValue = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            ElseIf sTestDataCellValue = "-" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            Else
                '### エラー詳細部＋エラー数格納 ###
                bIsExistCsvFile = False
                bIsExistTxtFile = False
                bIsExistHtmFile = False
                sErrDetail = ""
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName)
                    sTestLogFileName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    sTestLogBaseName = GetFileNameBase(sTestLogFileName)
                    sTestLogExt = GetFileNameExt(sTestLogFileName)
                    Select Case sTestLogExt
                        Case "csv"
                            '【チェック１】csv ファイル名が [モジュール名]_[項番].csv になっているか
                            bIsExistCsvFile = True
                            If sTestLogBaseName = sModuleName & "_" & sTestCaseNo Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "・csv ファイル名が命名規約と異なります！（[モジュール名]_[項番].csv）" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case "txt"
                            '【チェック２】txt ファイル名が [モジュール名].txt になっているか。
                            bIsExistTxtFile = True
                            If sTestLogBaseName = sModuleName Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "・txt ファイル名が命名規約と異なります！（[モジュール名].txt）" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case "htm"
                            '【チェック３】htm ファイル名が テスト結果報告書.htm になっているか。
                            bIsExistHtmFile = True
                            If sTestLogBaseName = "テスト結果報告書" Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "・htm ファイル名が命名規約と異なります！（テスト結果報告書.htm）" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case Else
                            '【チェック４】csv/txt/htm 以外のファイルが存在するか。
                            sErrDetail = sErrDetail & "・csv/txt/htm 以外のファイルが存在します！" & vbLf
                            gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End Select
                Next lTestLogNameIdx
                '【チェック５】htm、csv、txt が漏れなく存在するか
                If bIsExistCsvFile = True And _
                   bIsExistTxtFile = True And _
                   bIsExistHtmFile = True Then
                    'Do Nothing
                Else
                    If bIsExistCsvFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "・csv ファイルが記載されていません！" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                    If bIsExistTxtFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "・txt ファイルが記載されていません！" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                    If bIsExistHtmFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "・htm ファイルが記載されていません！" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                End If
                If Right(sErrDetail, 1) = vbLf Then
                    sErrDetail = Left(sErrDetail, Len(sErrDetail) - 1)
                Else
                    'Do Nothing
                End If
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
                
                '### チェック結果部格納 ###
                If sErrDetail = "" Then
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "OK!"
                Else
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "Error!"
                End If
            End If
            
            lErrMsgIdx = lErrMsgIdx + 1
            
        Next lTcInfoIdx
    Next lTcShtInfoIdx
End Function

Private Function StoreErrMsgExceptUT()
    Dim lTcShtInfoIdx As Long
    Dim lTcInfoIdx As Long
    Dim lTestLogNameIdx As Long
    Dim sTestLogFileName As String
    Dim sTestLogBaseName As String
    Dim sTestLogExt As String
    Dim sSheetName As String
    Dim sErrDetail As String
    Dim sTestCaseNo As String
    Dim sTestDataCellValue As String
    Dim lErrMsgIdx As Long
    
    lErrMsgIdx = 0
    For lTcShtInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo)
        sSheetName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName
        For lTcInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo)
            ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
            sTestCaseNo = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo
            
            '### シート名格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
            
            '### 項番格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
            
            '### 試験データ欄格納 ###
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
            
            If sTestDataCellValue = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            ElseIf sTestDataCellValue = "-" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            Else
                '### エラー詳細部＋エラー数格納 ###
                sErrDetail = ""
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName)
                    sTestLogFileName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    sTestLogBaseName = GetFileNameBase(sTestLogFileName)
                    sTestLogExt = GetFileNameExt(sTestLogFileName)
                    '【チェック１】ログファイル名に項番が含まれているか
                    If InStr(sTestLogBaseName, sTestCaseNo) > 0 Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "・ファイル名「" & sTestLogFileName & "」が命名規約と異なります！（[項番](_XXX).[拡張子]）" & vbLf
                        gtErrMsg.lWarningNum = gtErrMsg.lWarningNum + 1
                    End If
                Next lTestLogNameIdx
                If Right(sErrDetail, 1) = vbLf Then
                    sErrDetail = Left(sErrDetail, Len(sErrDetail) - 1)
                Else
                    'Do Nothing
                End If
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
                
                '### チェック結果部格納 ###
                If sErrDetail = "" Then
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "OK!"
                Else
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "Warning!"
                End If
            End If
            
            lErrMsgIdx = lErrMsgIdx + 1
            
        Next lTcInfoIdx
    Next lTcShtInfoIdx
End Function

Private Function OutpErrMsg( _
    ByRef wChkRsltBook As Workbook _
)
    Dim lOutpRowIdx As Long
    Dim lErrMsgRowIdx As Long
    Dim lOutpStrtRow As Long
    Dim lOutpEndRow As Long
    Dim lOutpStrtClm As Long
    Dim lOutpEndClm As Long
    
    '+++ シートコピー +++
    Call CopyRsltSht(wChkRsltBook, SHEET_NAME)
    
    With wChkRsltBook.Sheets(SHEET_NAME)
        lOutpStrtRow = ROW_ERR_MSG_STRT
        lOutpEndRow = lOutpStrtRow + UBound(gtErrMsg.atErrorMsgRow)
        lOutpStrtClm = CLM_STRT_PRE_CLM + 1
        lOutpEndClm = CLM_END_NEXT_CLM - 1
        
        '+++ セル書き込み +++
        '*** 項目書ファイル名 ***
        .Cells(ROW_TC_DOC_NAME, CLM_TC_DOC_NAME).Value = gtErrMsg.sDocFileName
        
        '*** エラー数 ***
        .Cells(ROW_ERR_NUM, CLM_TC_ERR_NUM).Value = gtErrMsg.lErrNum
        
        '*** ワーニング数 ***
        .Cells(ROW_WARN_NUM, CLM_TC_WARN_NUM).Value = gtErrMsg.lWarningNum
        
        '*** エラー内容 ***
        For lErrMsgRowIdx = 0 To UBound(gtErrMsg.atErrorMsgRow)
            lOutpRowIdx = lOutpStrtRow + lErrMsgRowIdx
            '=== シート名 ===
            .Cells(lOutpRowIdx, CLM_SHT_NAME).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sSheetName
            '=== 項番 ===
            .Cells(lOutpRowIdx, CLM_TC_NO).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcNo
            '=== 試験データ ===
            .Cells(lOutpRowIdx, CLM_TC_DATA).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcData
            '=== チェック結果 ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== エラー詳細 ===
            .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
        Next lErrMsgRowIdx
        
        '=== 書式コピー＆ペースト ===
        .Range( _
                .Cells(lOutpStrtRow, lOutpStrtClm), _
                .Cells(lOutpStrtRow, lOutpEndClm) _
              ).Copy
        .Range( _
                .Cells(lOutpStrtRow, lOutpStrtClm), _
                .Cells(lOutpEndRow, lOutpEndClm) _
              ).PasteSpecial (xlPasteFormats)
        Application.CutCopyMode = False
        
        '=== オートフィルタ追加 ===
        .Range( _
                .Cells(lOutpStrtRow - 1, lOutpStrtClm), _
                .Cells(lOutpEndRow, lOutpEndClm) _
              ).AutoFilter
        
        .Cells(1, 1).Select
    End With
End Function

