Attribute VB_Name = "Chk_A_TestData"
Option Explicit

Private Const SHEET_NAME = "(A)"

Private Enum E_ROW
    ROW_TC_DOC_NAME = 6
    ROW_RESERVE_01 '���g�p
    ROW_ERR_NUM
    ROW_WARN_NUM
    ROW_RESERVE_02 '���g�p
    ROW_ERR_MSG_TITLE
    ROW_ERR_MSG_STRT
End Enum

Private Enum E_CLM
    CLM_TC_DOC_NAME = 4
    CLM_TC_ERR_NUM = 4
    CLM_TC_WARN_NUM = 4
    
    CLM_STRT_PRE_CLM = 2 '�őO�� - 1
    CLM_SHT_NAME
    CLM_TC_NO
    CLM_TC_DATA
    CLM_RESERVE_01 '���g�p
    CLM_CHK_RSLT
    CLM_CHK_DETAIL
    CLM_END_NEXT_CLM '�Ō�� + 1
End Enum

Private Type T_ERROR_MSG_ROW
    sBookName As String  '���󖢎g�p
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
    '### �^�C�g�����i�[ ###
    gtErrMsg.sDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    'gtErrMsg.lErrNum '�G���[���̓G���[���b�Z�[�W���i�[�����ɂĊi�[����
    
    '### �G���[���b�Z�[�W���i�[ ###
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
        Call StoreErrMsgUT
    Else
        Call StoreErrMsgExceptUT
    End If
    
    '���ʂ܂Ƃ߃V�[�g�փG���[��/���[�j���O����ʒm
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
            
            '### �V�[�g���i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
            
            '### ���Ԋi�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
            
            '### �����f�[�^���i�[ ###
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
            
            If sTestDataCellValue = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            ElseIf sTestDataCellValue = "-" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            Else
                '### �G���[�ڍו��{�G���[���i�[ ###
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
                            '�y�`�F�b�N�P�zcsv �t�@�C������ [���W���[����]_[����].csv �ɂȂ��Ă��邩
                            bIsExistCsvFile = True
                            If sTestLogBaseName = sModuleName & "_" & sTestCaseNo Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "�Ecsv �t�@�C�����������K��ƈقȂ�܂��I�i[���W���[����]_[����].csv�j" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case "txt"
                            '�y�`�F�b�N�Q�ztxt �t�@�C������ [���W���[����].txt �ɂȂ��Ă��邩�B
                            bIsExistTxtFile = True
                            If sTestLogBaseName = sModuleName Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "�Etxt �t�@�C�����������K��ƈقȂ�܂��I�i[���W���[����].txt�j" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case "htm"
                            '�y�`�F�b�N�R�zhtm �t�@�C������ �e�X�g���ʕ񍐏�.htm �ɂȂ��Ă��邩�B
                            bIsExistHtmFile = True
                            If sTestLogBaseName = "�e�X�g���ʕ񍐏�" Then
                                'Do Nothing
                            Else
                                sErrDetail = sErrDetail & "�Ehtm �t�@�C�����������K��ƈقȂ�܂��I�i�e�X�g���ʕ񍐏�.htm�j" & vbLf
                                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                            End If
                        Case Else
                            '�y�`�F�b�N�S�zcsv/txt/htm �ȊO�̃t�@�C�������݂��邩�B
                            sErrDetail = sErrDetail & "�Ecsv/txt/htm �ȊO�̃t�@�C�������݂��܂��I" & vbLf
                            gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End Select
                Next lTestLogNameIdx
                '�y�`�F�b�N�T�zhtm�Acsv�Atxt ���R��Ȃ����݂��邩
                If bIsExistCsvFile = True And _
                   bIsExistTxtFile = True And _
                   bIsExistHtmFile = True Then
                    'Do Nothing
                Else
                    If bIsExistCsvFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "�Ecsv �t�@�C�����L�ڂ���Ă��܂���I" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                    If bIsExistTxtFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "�Etxt �t�@�C�����L�ڂ���Ă��܂���I" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                    If bIsExistHtmFile = True Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "�Ehtm �t�@�C�����L�ڂ���Ă��܂���I" & vbLf
                        gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                    End If
                End If
                If Right(sErrDetail, 1) = vbLf Then
                    sErrDetail = Left(sErrDetail, Len(sErrDetail) - 1)
                Else
                    'Do Nothing
                End If
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
                
                '### �`�F�b�N���ʕ��i�[ ###
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
            
            '### �V�[�g���i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
            
            '### ���Ԋi�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
            
            '### �����f�[�^���i�[ ###
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
            
            If sTestDataCellValue = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            ElseIf sTestDataCellValue = "-" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            Else
                '### �G���[�ڍו��{�G���[���i�[ ###
                sErrDetail = ""
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName)
                    sTestLogFileName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    sTestLogBaseName = GetFileNameBase(sTestLogFileName)
                    sTestLogExt = GetFileNameExt(sTestLogFileName)
                    '�y�`�F�b�N�P�z���O�t�@�C�����ɍ��Ԃ��܂܂�Ă��邩
                    If InStr(sTestLogBaseName, sTestCaseNo) > 0 Then
                        'Do Nothing
                    Else
                        sErrDetail = sErrDetail & "�E�t�@�C�����u" & sTestLogFileName & "�v�������K��ƈقȂ�܂��I�i[����](_XXX).[�g���q]�j" & vbLf
                        gtErrMsg.lWarningNum = gtErrMsg.lWarningNum + 1
                    End If
                Next lTestLogNameIdx
                If Right(sErrDetail, 1) = vbLf Then
                    sErrDetail = Left(sErrDetail, Len(sErrDetail) - 1)
                Else
                    'Do Nothing
                End If
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
                
                '### �`�F�b�N���ʕ��i�[ ###
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
    
    '+++ �V�[�g�R�s�[ +++
    Call CopyRsltSht(wChkRsltBook, SHEET_NAME)
    
    With wChkRsltBook.Sheets(SHEET_NAME)
        lOutpStrtRow = ROW_ERR_MSG_STRT
        lOutpEndRow = lOutpStrtRow + UBound(gtErrMsg.atErrorMsgRow)
        lOutpStrtClm = CLM_STRT_PRE_CLM + 1
        lOutpEndClm = CLM_END_NEXT_CLM - 1
        
        '+++ �Z���������� +++
        '*** ���ڏ��t�@�C���� ***
        .Cells(ROW_TC_DOC_NAME, CLM_TC_DOC_NAME).Value = gtErrMsg.sDocFileName
        
        '*** �G���[�� ***
        .Cells(ROW_ERR_NUM, CLM_TC_ERR_NUM).Value = gtErrMsg.lErrNum
        
        '*** ���[�j���O�� ***
        .Cells(ROW_WARN_NUM, CLM_TC_WARN_NUM).Value = gtErrMsg.lWarningNum
        
        '*** �G���[���e ***
        For lErrMsgRowIdx = 0 To UBound(gtErrMsg.atErrorMsgRow)
            lOutpRowIdx = lOutpStrtRow + lErrMsgRowIdx
            '=== �V�[�g�� ===
            .Cells(lOutpRowIdx, CLM_SHT_NAME).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sSheetName
            '=== ���� ===
            .Cells(lOutpRowIdx, CLM_TC_NO).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcNo
            '=== �����f�[�^ ===
            .Cells(lOutpRowIdx, CLM_TC_DATA).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcData
            '=== �`�F�b�N���� ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== �G���[�ڍ� ===
            .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
        Next lErrMsgRowIdx
        
        '=== �����R�s�[���y�[�X�g ===
        .Range( _
                .Cells(lOutpStrtRow, lOutpStrtClm), _
                .Cells(lOutpStrtRow, lOutpEndClm) _
              ).Copy
        .Range( _
                .Cells(lOutpStrtRow, lOutpStrtClm), _
                .Cells(lOutpEndRow, lOutpEndClm) _
              ).PasteSpecial (xlPasteFormats)
        Application.CutCopyMode = False
        
        '=== �I�[�g�t�B���^�ǉ� ===
        .Range( _
                .Cells(lOutpStrtRow - 1, lOutpStrtClm), _
                .Cells(lOutpEndRow, lOutpEndClm) _
              ).AutoFilter
        
        .Cells(1, 1).Select
    End With
End Function

