Attribute VB_Name = "Chk_C_TestDataOmission"
Option Explicit

Private Const SHEET_NAME = "(C)"

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
    CLM_ATTR
    CLM_PATH
    CLM_FILE_NAME
    CLM_RESERVE_01 '未使用
    CLM_CHK_RSLT
    CLM_ERR_DETAIL
    CLM_END_NEXT_CLM '最後列 + 1
End Enum

Private Type T_ERROR_MSG_ROW
    sPathType As String
    sPath As String
    sFileName As String
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

Public Function ChkTestDataOmissionInit()
    Dim tErrMsg As T_ERROR_MSG
    gtErrMsg = tErrMsg
End Function

Public Function ChkTestDataOmissionMain( _
    ByRef wChkRsltBook As Workbook _
)
    Call StoreErrMsg
    Call OutpErrMsg(wChkRsltBook)
End Function

Private Function StoreErrMsg()
    Dim lPathListIdx As Long
    Dim sPathListPath As String
    Dim sRelativePath As String
    Dim ePathType As E_PATH_TYPE
    Dim sErrDetail As String
    Dim lErrMsgIdx As Long
    Dim sFileExt As String
    Dim sLogDirPathBeforeDocName As String
    
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
    For lPathListIdx = 0 To UBound(gatPathList)
        ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
        sPathListPath = gatPathList(lPathListIdx).sPath
        sRelativePath = Replace(sPathListPath, gtErrMsg.sLogDirPath & "\", "")
        ePathType = gatPathList(lPathListIdx).ePathType
        If ePathType = PATH_TYPE_DIRECTORY Then
            '### 属性格納 ###
            If ePathType = PATH_TYPE_DIRECTORY Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "Directory"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "File"
            End If
            '### パス格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPath = sRelativePath
            '### ファイル名格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sFileName = "-"
            '### チェック結果格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            '### エラー詳細格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
        Else
            '### 属性格納 ###
            If ePathType = PATH_TYPE_DIRECTORY Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "Directory"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "File"
            End If
            '### パス格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPath = sRelativePath
            '### ファイル名格納 ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sFileName = GetFileName(sRelativePath)
            '### エラー詳細格納 ###
            If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
                sFileExt = GetFileNameExt(sPathListPath)
                Select Case sFileExt
                    Case "csv"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            sErrDetail = "・このログファイルは試験データ欄に記載されていません！"
                        End If
                    Case "htm"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            sErrDetail = "・このログファイルは試験データ欄に記載されていません！"
                        End If
                    Case "txt"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            If GetStrNum(sRelativePath, "\") = 4 Then '「項目書名\シート名\TestCoverLog\ソースファイル名\」配下に格納されているか
                                sErrDetail = ""
                            Else
                                sErrDetail = "・このフォルダには txt ファイルを格納してはいけません！"
                            End If
                        End If
                    Case Else
                        sErrDetail = "・csv/txt/htm ファイル以外のファイルが格納されています！"
                End Select
            Else
                If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                    sErrDetail = ""
                Else
                    sErrDetail = "・このログファイルは試験データ欄に記載されていません！"
                End If
            End If
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
            
            '### チェック結果格納 ###
            If sErrDetail = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "OK!"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "Error!"
                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
            End If
        End If
        
        lErrMsgIdx = lErrMsgIdx + 1
    Next lPathListIdx
    
    '結果まとめシートへエラー数/ワーニング数を通知
    Call SetErrNum2Summary(FUNC_TYPE_C, gtErrMsg.lErrNum, gtErrMsg.lWarningNum)
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
        
        '*** 項目書ファイル名 ***
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtErrMsg.sLogDirPath
        
        '*** エラー数 ***
        .Cells(ROW_ERR_NUM, CLM_TC_ERR_NUM).Value = gtErrMsg.lErrNum
        
        '*** ワーニング数 ***
        .Cells(ROW_WARN_NUM, CLM_TC_WARN_NUM).Value = gtErrMsg.lWarningNum
        
        '*** エラー内容 ***
        For lErrMsgRowIdx = 0 To UBound(gtErrMsg.atErrorMsgRow)
            lOutpRowIdx = lOutpStrtRow + lErrMsgRowIdx
            '=== 属性 ===
            .Cells(lOutpRowIdx, CLM_ATTR).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sPathType
            '=== ファイルパス ===
            .Cells(lOutpRowIdx, CLM_PATH).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sPath
            '=== ファイル名 ===
            .Cells(lOutpRowIdx, CLM_FILE_NAME).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sFileName
            '=== チェック結果 ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== エラー詳細 ===
            .Cells(lOutpRowIdx, CLM_ERR_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
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


