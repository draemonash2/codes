Attribute VB_Name = "Chk_B_ExistLogFile"
Option Explicit

Private Const SHEET_NAME = "(B)"

Private Enum E_ROW
    ROW_TC_DOC_NAME = 6
    ROW_LOG_DIR_PATH
    ROW_RESERVE_01 '未使用
    ROW_ERR_NUM
    ROW_WARN_NUM
    ROW_RESERVE_02 '未使用
    ROW_ERR_MSG_TITLE
    ROW_ERR_MSG_STRT
End Enum

Private Enum E_CLM
    CLM_TC_DOC_NAME = 4
    CLM_LOG_DIR_PATH = 4
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
    sLogDirPath As String
    lErrNum As Long
    lWarningNum As Long
    atErrorMsgRow() As T_ERROR_MSG_ROW
End Type

Private gtErrMsg As T_ERROR_MSG

Public Function ChkExistLogFileInit()
    Dim tErrMsg As T_ERROR_MSG
    gtErrMsg = tErrMsg
End Function

Public Function ChkExistLogFileMain( _
    ByRef wChkRsltBook As Workbook _
)
    Call StoreErrMsg
    Call OutpErrMsg(wChkRsltBook)
End Function

Private Function StoreErrMsg()
    Dim lTcShtInfoIdx As Long
    Dim lTcInfoIdx As Long
    Dim lTestLogNameIdx As Long
    Dim sTestLogFileName As String
    Dim sTestLogBaseName As String
    Dim sTestLogExt As String
    Dim sSheetName As String
    Dim sModuleName As String
    Dim sTestCaseNo As String
    Dim sTestDataCellValue As String
    Dim lErrMsgIdx As Long
    Dim sExpLogFileRelativePath As String
    Dim sExpLogFilePath As String
    Dim bisChkExec As Boolean
    
    '### タイトル部格納 ###
    gtErrMsg.sDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        gtErrMsg.sLogDirPath = _
            gtInputInfo.sTestLogDirPath
    Else
        gtErrMsg.sLogDirPath = _
            gtInputInfo.sTestLogDirPath & "\" & _
            gtInputInfo.sSubjectName
    End If
    'gtErrMsg.lErrNum 'エラー数はエラーメッセージ部格納処理にて格納する
    
    '### エラーメッセージ部格納 ###
    lErrMsgIdx = 0
    For lTcShtInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo)
        sModuleName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName
        sSheetName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName
        For lTcInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo)
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            sTestCaseNo = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo
            If sTestDataCellValue = "" Or sTestDataCellValue = "-" Then
                ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
                
                '### シート名格納 ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
                
                '### 項番格納 ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
                
                '### 試験データ欄格納 ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
                
                '### エラー詳細部＋エラー数格納 ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
                
                '### チェック結果部格納 ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
                
                lErrMsgIdx = lErrMsgIdx + 1
            Else
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName)
                    ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
                    sTestLogFileName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    sTestLogBaseName = GetFileNameBase(sTestLogFileName)
                    sTestLogExt = GetFileNameExt(sTestLogFileName)
                    
                    '### シート名格納 ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
                    
                    '### 項番格納 ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
                    
                    '### 試験データ欄格納 ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestLogFileName
                    
                    '### エラー詳細部＋エラー数格納 ###
                    '【チェック１】テストデータに記載のログファイルが存在するか
                    If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
                        Select Case sTestLogExt
                            Case "csv"
                                bisChkExec = True
                                sExpLogFileRelativePath = _
                                    GetFileNameBase(gtInputInfo.sTestDocFilePath) & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                            Case "txt"
                                bisChkExec = True
                                sExpLogFileRelativePath = _
                                    GetFileNameBase(gtInputInfo.sTestDocFilePath) & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                    "TestCoverLog\" & _
                                    GetFileName(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sSrcFileName) & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                            Case "htm"
                                bisChkExec = True
                                sExpLogFileRelativePath = _
                                    GetFileNameBase(gtInputInfo.sTestDocFilePath) & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                            Case Else
                                'csv/txt/htm 以外のファイルは期待するファイルパスが分からないため、チェックしない
                                bisChkExec = False
                                sExpLogFileRelativePath = ""
                        End Select
                    Else
                        bisChkExec = True
                        sExpLogFileRelativePath = _
                            GetFileNameBase(gtInputInfo.sTestDocFilePath) & "\" & _
                            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    End If
                    sExpLogFilePath = gtErrMsg.sLogDirPath & "\" & sExpLogFileRelativePath
                    
                    If bisChkExec = True Then
                        If ChkFileExist(sExpLogFilePath) = False Then
                            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = _
                                "・ファイルが規定フォルダ内に存在しません！" & vbNewLine & _
                                "【期待ファイルパス】" & sExpLogFileRelativePath
                            gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                        Else
                            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                        End If
                    Else
                        gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                    End If
                    
                    '### チェック結果部格納 ###
                    If gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = "" Then
                        gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "OK!"
                    Else
                        gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "Error!"
                    End If
                    lErrMsgIdx = lErrMsgIdx + 1
                Next lTestLogNameIdx
            End If
        Next lTcInfoIdx
    Next lTcShtInfoIdx
    
    '結果まとめシートへエラー数/ワーニング数を通知
    Call SetErrNum2Summary(FUNC_TYPE_B, gtErrMsg.lErrNum, gtErrMsg.lWarningNum)
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
    Dim sCommentTxt As String
    Dim sTestDataFileName As String
    
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
        
        '*** ログフォルダパス ***
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtErrMsg.sLogDirPath
        
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
            sTestDataFileName = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcData
            .Cells(lOutpRowIdx, CLM_TC_DATA).Value = sTestDataFileName
            '=== チェック結果 ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== エラー詳細 ===
            .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
            '=== コメント付与 ===
            If gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt = "OK!" Or _
               gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt = "-" Then
                'Do Nothing
            Else
                Select Case gtInputInfo.eTrgtPhase
                    Case TRGT_PHASE_UT
                        sCommentTxt = _
                                        "ログファイルが格納されているか確認してください。" & vbNewLine & _
                                        "ファイルが存在する場合は、フォルダ構成を見直してください。" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "～規定フォルダ構成～" & vbNewLine & _
                                        "  XX_単体試験ログ" & vbNewLine & _
                                        "    └ [案件名]" & vbNewLine & _
                                        "      └ [項目書ファイル名（拡張子なし）]" & vbNewLine & _
                                        "        └ [シート名]" & vbNewLine & _
                                        "          ├ [モジュール名]_[項番].csv" & vbNewLine & _
                                        "          ├ テスト結果報告書.htm" & vbNewLine & _
                                        "          └ TestCoverLog" & vbNewLine & _
                                        "            └ [ソースファイル名（拡張子あり）]" & vbNewLine & _
                                        "              └ [モジュール名].txt"
                    Case TRGT_PHASE_CT
                        sCommentTxt = _
                                        "ログファイルが格納されているか確認してください。" & vbNewLine & _
                                        "ファイルが存在する場合は、フォルダ構成を見直してください。" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "～規定フォルダ構成～" & vbNewLine & _
                                        "  XX_結合試験ログ" & vbNewLine & _
                                        "    └ [案件名]" & vbNewLine & _
                                        "      └ [項目書ファイル名（拡張子なし）]" & vbNewLine & _
                                        "        └ [シート名]" & vbNewLine & _
                                        "          └ [項番].[拡張子]"
                    Case TRGT_PHASE_FT
                        sCommentTxt = _
                                        "ログファイルが格納されているか確認してください。" & vbNewLine & _
                                        "ファイルが存在する場合は、フォルダ構成を見直してください。" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "～規定フォルダ構成～" & vbNewLine & _
                                        "  XX_機能試験ログ" & vbNewLine & _
                                        "    └ [案件名]" & vbNewLine & _
                                        "      └ [項目書ファイル名（拡張子なし）]" & vbNewLine & _
                                        "        └ [シート名]" & vbNewLine & _
                                        "          └ [項番].[拡張子]"
                    Case TRGT_PHASE_ST
                        sCommentTxt = _
                                        "ログファイルが格納されているか確認してください。" & vbNewLine & _
                                        "ファイルが存在する場合は、フォルダ構成を見直してください。" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "～規定フォルダ構成～" & vbNewLine & _
                                        "  XX_システム試験ログ" & vbNewLine & _
                                        "    └ [項目書ファイル名（拡張子なし）]" & vbNewLine & _
                                        "      └ [シート名]" & vbNewLine & _
                                        "        └ [項番].[拡張子]"
                    Case Else
                        Stop
                End Select
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).AddComment
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Visible = True
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Text Text:=sCommentTxt
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Shape.TextFrame.AutoSize = True
            End If
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

