Attribute VB_Name = "Chk_B_ExistLogFile"
Option Explicit

Private Const SHEET_NAME = "(B)"

Private Enum E_ROW
    ROW_TC_DOC_NAME = 6
    ROW_LOG_DIR_PATH
    ROW_RESERVE_01 '���g�p
    ROW_ERR_NUM
    ROW_WARN_NUM
    ROW_RESERVE_02 '���g�p
    ROW_ERR_MSG_TITLE
    ROW_ERR_MSG_STRT
End Enum

Private Enum E_CLM
    CLM_TC_DOC_NAME = 4
    CLM_LOG_DIR_PATH = 4
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
    
    '### �^�C�g�����i�[ ###
    gtErrMsg.sDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        gtErrMsg.sLogDirPath = _
            gtInputInfo.sTestLogDirPath
    Else
        gtErrMsg.sLogDirPath = _
            gtInputInfo.sTestLogDirPath & "\" & _
            gtInputInfo.sSubjectName
    End If
    'gtErrMsg.lErrNum '�G���[���̓G���[���b�Z�[�W���i�[�����ɂĊi�[����
    
    '### �G���[���b�Z�[�W���i�[ ###
    lErrMsgIdx = 0
    For lTcShtInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo)
        sModuleName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName
        sSheetName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName
        For lTcInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo)
            sTestDataCellValue = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue
            sTestCaseNo = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo
            If sTestDataCellValue = "" Or sTestDataCellValue = "-" Then
                ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
                
                '### �V�[�g���i�[ ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
                
                '### ���Ԋi�[ ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
                
                '### �����f�[�^���i�[ ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestDataCellValue
                
                '### �G���[�ڍו��{�G���[���i�[ ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
                
                '### �`�F�b�N���ʕ��i�[ ###
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
                
                lErrMsgIdx = lErrMsgIdx + 1
            Else
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName)
                    ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
                    sTestLogFileName = gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                    sTestLogBaseName = GetFileNameBase(sTestLogFileName)
                    sTestLogExt = GetFileNameExt(sTestLogFileName)
                    
                    '### �V�[�g���i�[ ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sSheetName = sSheetName
                    
                    '### ���Ԋi�[ ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcNo = sTestCaseNo
                    
                    '### �����f�[�^���i�[ ###
                    gtErrMsg.atErrorMsgRow(lErrMsgIdx).sTcData = sTestLogFileName
                    
                    '### �G���[�ڍו��{�G���[���i�[ ###
                    '�y�`�F�b�N�P�z�e�X�g�f�[�^�ɋL�ڂ̃��O�t�@�C�������݂��邩
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
                                'csv/txt/htm �ȊO�̃t�@�C���͊��҂���t�@�C���p�X��������Ȃ����߁A�`�F�b�N���Ȃ�
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
                                "�E�t�@�C�����K��t�H���_���ɑ��݂��܂���I" & vbNewLine & _
                                "�y���҃t�@�C���p�X�z" & sExpLogFileRelativePath
                            gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
                        Else
                            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                        End If
                    Else
                        gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
                    End If
                    
                    '### �`�F�b�N���ʕ��i�[ ###
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
    
    '���ʂ܂Ƃ߃V�[�g�փG���[��/���[�j���O����ʒm
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
        
        '*** ���O�t�H���_�p�X ***
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtErrMsg.sLogDirPath
        
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
            sTestDataFileName = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sTcData
            .Cells(lOutpRowIdx, CLM_TC_DATA).Value = sTestDataFileName
            '=== �`�F�b�N���� ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== �G���[�ڍ� ===
            .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
            '=== �R�����g�t�^ ===
            If gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt = "OK!" Or _
               gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt = "-" Then
                'Do Nothing
            Else
                Select Case gtInputInfo.eTrgtPhase
                    Case TRGT_PHASE_UT
                        sCommentTxt = _
                                        "���O�t�@�C�����i�[����Ă��邩�m�F���Ă��������B" & vbNewLine & _
                                        "�t�@�C�������݂���ꍇ�́A�t�H���_�\�����������Ă��������B" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "�`�K��t�H���_�\���`" & vbNewLine & _
                                        "  XX_�P�̎������O" & vbNewLine & _
                                        "    �� [�Č���]" & vbNewLine & _
                                        "      �� [���ڏ��t�@�C�����i�g���q�Ȃ��j]" & vbNewLine & _
                                        "        �� [�V�[�g��]" & vbNewLine & _
                                        "          �� [���W���[����]_[����].csv" & vbNewLine & _
                                        "          �� �e�X�g���ʕ񍐏�.htm" & vbNewLine & _
                                        "          �� TestCoverLog" & vbNewLine & _
                                        "            �� [�\�[�X�t�@�C�����i�g���q����j]" & vbNewLine & _
                                        "              �� [���W���[����].txt"
                    Case TRGT_PHASE_CT
                        sCommentTxt = _
                                        "���O�t�@�C�����i�[����Ă��邩�m�F���Ă��������B" & vbNewLine & _
                                        "�t�@�C�������݂���ꍇ�́A�t�H���_�\�����������Ă��������B" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "�`�K��t�H���_�\���`" & vbNewLine & _
                                        "  XX_�����������O" & vbNewLine & _
                                        "    �� [�Č���]" & vbNewLine & _
                                        "      �� [���ڏ��t�@�C�����i�g���q�Ȃ��j]" & vbNewLine & _
                                        "        �� [�V�[�g��]" & vbNewLine & _
                                        "          �� [����].[�g���q]"
                    Case TRGT_PHASE_FT
                        sCommentTxt = _
                                        "���O�t�@�C�����i�[����Ă��邩�m�F���Ă��������B" & vbNewLine & _
                                        "�t�@�C�������݂���ꍇ�́A�t�H���_�\�����������Ă��������B" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "�`�K��t�H���_�\���`" & vbNewLine & _
                                        "  XX_�@�\�������O" & vbNewLine & _
                                        "    �� [�Č���]" & vbNewLine & _
                                        "      �� [���ڏ��t�@�C�����i�g���q�Ȃ��j]" & vbNewLine & _
                                        "        �� [�V�[�g��]" & vbNewLine & _
                                        "          �� [����].[�g���q]"
                    Case TRGT_PHASE_ST
                        sCommentTxt = _
                                        "���O�t�@�C�����i�[����Ă��邩�m�F���Ă��������B" & vbNewLine & _
                                        "�t�@�C�������݂���ꍇ�́A�t�H���_�\�����������Ă��������B" & vbNewLine & _
                                        "" & vbNewLine & _
                                        "�`�K��t�H���_�\���`" & vbNewLine & _
                                        "  XX_�V�X�e���������O" & vbNewLine & _
                                        "    �� [���ڏ��t�@�C�����i�g���q�Ȃ��j]" & vbNewLine & _
                                        "      �� [�V�[�g��]" & vbNewLine & _
                                        "        �� [����].[�g���q]"
                    Case Else
                        Stop
                End Select
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).AddComment
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Visible = True
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Text Text:=sCommentTxt
                .Cells(lOutpRowIdx, CLM_CHK_DETAIL).Comment.Shape.TextFrame.AutoSize = True
            End If
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

