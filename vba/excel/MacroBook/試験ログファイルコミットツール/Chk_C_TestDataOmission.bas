Attribute VB_Name = "Chk_C_TestDataOmission"
Option Explicit

Private Const SHEET_NAME = "(C)"

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
    CLM_ATTR
    CLM_PATH
    CLM_FILE_NAME
    CLM_RESERVE_01 '���g�p
    CLM_CHK_RSLT
    CLM_ERR_DETAIL
    CLM_END_NEXT_CLM '�Ō�� + 1
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
    For lPathListIdx = 0 To UBound(gatPathList)
        ReDim Preserve gtErrMsg.atErrorMsgRow(lErrMsgIdx)
        sPathListPath = gatPathList(lPathListIdx).sPath
        sRelativePath = Replace(sPathListPath, gtErrMsg.sLogDirPath & "\", "")
        ePathType = gatPathList(lPathListIdx).ePathType
        If ePathType = PATH_TYPE_DIRECTORY Then
            '### �����i�[ ###
            If ePathType = PATH_TYPE_DIRECTORY Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "Directory"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "File"
            End If
            '### �p�X�i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPath = sRelativePath
            '### �t�@�C�����i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sFileName = "-"
            '### �`�F�b�N���ʊi�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "-"
            '### �G���[�ڍ׊i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = ""
        Else
            '### �����i�[ ###
            If ePathType = PATH_TYPE_DIRECTORY Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "Directory"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPathType = "File"
            End If
            '### �p�X�i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sPath = sRelativePath
            '### �t�@�C�����i�[ ###
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sFileName = GetFileName(sRelativePath)
            '### �G���[�ڍ׊i�[ ###
            If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
                sFileExt = GetFileNameExt(sPathListPath)
                Select Case sFileExt
                    Case "csv"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            sErrDetail = "�E���̃��O�t�@�C���͎����f�[�^���ɋL�ڂ���Ă��܂���I"
                        End If
                    Case "htm"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            sErrDetail = "�E���̃��O�t�@�C���͎����f�[�^���ɋL�ڂ���Ă��܂���I"
                        End If
                    Case "txt"
                        If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                            sErrDetail = ""
                        Else
                            If GetStrNum(sRelativePath, "\") = 4 Then '�u���ڏ���\�V�[�g��\TestCoverLog\�\�[�X�t�@�C����\�v�z���Ɋi�[����Ă��邩
                                sErrDetail = ""
                            Else
                                sErrDetail = "�E���̃t�H���_�ɂ� txt �t�@�C�����i�[���Ă͂����܂���I"
                            End If
                        End If
                    Case Else
                        sErrDetail = "�Ecsv/txt/htm �t�@�C���ȊO�̃t�@�C�����i�[����Ă��܂��I"
                End Select
            Else
                If gtTestDocInfo.oLogExpPathList.exists(sRelativePath) = True Then
                    sErrDetail = ""
                Else
                    sErrDetail = "�E���̃��O�t�@�C���͎����f�[�^���ɋL�ڂ���Ă��܂���I"
                End If
            End If
            gtErrMsg.atErrorMsgRow(lErrMsgIdx).sErrDetail = sErrDetail
            
            '### �`�F�b�N���ʊi�[ ###
            If sErrDetail = "" Then
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "OK!"
            Else
                gtErrMsg.atErrorMsgRow(lErrMsgIdx).sChkRslt = "Error!"
                gtErrMsg.lErrNum = gtErrMsg.lErrNum + 1
            End If
        End If
        
        lErrMsgIdx = lErrMsgIdx + 1
    Next lPathListIdx
    
    '���ʂ܂Ƃ߃V�[�g�փG���[��/���[�j���O����ʒm
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
        
        '*** ���ڏ��t�@�C���� ***
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtErrMsg.sLogDirPath
        
        '*** �G���[�� ***
        .Cells(ROW_ERR_NUM, CLM_TC_ERR_NUM).Value = gtErrMsg.lErrNum
        
        '*** ���[�j���O�� ***
        .Cells(ROW_WARN_NUM, CLM_TC_WARN_NUM).Value = gtErrMsg.lWarningNum
        
        '*** �G���[���e ***
        For lErrMsgRowIdx = 0 To UBound(gtErrMsg.atErrorMsgRow)
            lOutpRowIdx = lOutpStrtRow + lErrMsgRowIdx
            '=== ���� ===
            .Cells(lOutpRowIdx, CLM_ATTR).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sPathType
            '=== �t�@�C���p�X ===
            .Cells(lOutpRowIdx, CLM_PATH).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sPath
            '=== �t�@�C���� ===
            .Cells(lOutpRowIdx, CLM_FILE_NAME).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sFileName
            '=== �`�F�b�N���� ===
            .Cells(lOutpRowIdx, CLM_CHK_RSLT).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sChkRslt
            '=== �G���[�ڍ� ===
            .Cells(lOutpRowIdx, CLM_ERR_DETAIL).Value = gtErrMsg.atErrorMsgRow(lErrMsgRowIdx).sErrDetail
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


