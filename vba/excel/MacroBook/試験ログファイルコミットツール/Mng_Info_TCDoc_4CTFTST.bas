Attribute VB_Name = "Mng_Info_TCDoc_4CTFTST"
Option Explicit

Private Const TC_SHT_KEY_WORD = "��������"
Private Const SRCH_KEYWORD_TC_TESTER = "�]����"
Private Const SRCH_KEYWORD_TEST_RSLT = "����"
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
        '### �����t�F�[�Y�擾 ###
        gtTestDocInfo.eTrgtPhase = gtInputInfo.eTrgtPhase
        
        '### �u�b�N���擾 ###
        gtTestDocInfo.sTcDocName = wTrgtBook.Name
        
        For lShtIdx = 1 To wTrgtBook.Sheets.Count
            '���ڃV�[�g����
            If InStr(wTrgtBook.Sheets(lShtIdx).Name, TC_SHT_KEY_WORD) > 0 Then
                If Sgn(gtTestDocInfo.atTcShtInfo) = 0 Then
                    lTcShtInfoIdx = 0
                Else
                    lTcShtInfoIdx = UBound(gtTestDocInfo.atTcShtInfo) + 1
                End If
                ReDim Preserve gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx)
                
                '==================
                '=== �Z���l�擾 ===
                '==================
                '### �V�[�g���擾 ###
                gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName = wTrgtBook.Sheets(lShtIdx).Name
                
                '====================
                '=== �Z�����擾 ===
                '====================
                '### ���� �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TC_NO
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTcNoClm = tNearCellData.lClm
                lTcStrtRow = tNearCellData.lRow + 2
                lTcEndRow = .Sheets(lShtIdx).Cells(.Sheets(lShtIdx).Rows.Count, tTcClmInfo.lTcNoClm).End(xlUp).Row
                lTcNum = lTcEndRow - lTcStrtRow + 1
                
                '### �]���� �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TC_TESTER
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTesterClm = tNearCellData.lClm
                
                '### ���� �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TEST_RSLT
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTestResultClm = tNearCellData.lClm
                
                '### �N���� �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TEST_DATE
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTestDateClm = tNearCellData.lClm
                
                '### Rev�iHEX/ABS�j �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_REV_HEXABS
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTestRevHexAbsClm = tNearCellData.lClm
                
                '### Rev�iA2l�j �Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_REV_A2L
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTestRevA2LClm = tNearCellData.lClm
                
                '### �����f�[�^�Z�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TEST_DATA
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 0)
                If tNearCellData.bIsCellDataExist = True Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & .Name & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                tTcClmInfo.lTcDataClm = tNearCellData.lClm
                
                '�G���[�o��
                Call OutpErrorMsg(ERROR_PROC_STOP)
                
                '### ���ڏ��擾 ###
                '���ڐ��O
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
                
                '�P�̎����p�̏��ɉ����i�[����Ă��Ȃ����Ƃ��m�F
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
        '���Ԃ��u-�v���󗓂̏ꍇ�A����
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
            
            '### ���Ԏ擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo = sTcNo
            
            '### �]���Ҏ擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTester = sTester
            
            '### �N�����擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDate = sTestDate
            
            '### ����擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestResult = sTestResult
            
            '### Rev�iHEX/ABS�j�擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevHexAbs = sTestRevHexAbs
            
            '### Rev�iA2L�j�擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevA2L = sTestRevA2L
            
            '### �����f�[�^�{���҂w�w�w�t�@�C���p�X���X�g�擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDataCellValue = sTestDataCellValue
            If sTestDataCellValue = "" Then
                'Do Nothing
            Else
                '�Z�����̉��s���f���~�^�Ƃ��Ĕz��ɕ���
                asTestDataFileNames = GetTestDataArray(sTestDataCellValue)
                
                ReDim Preserve gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(UBound(asTestDataFileNames))
                For lTestDataIdx = 0 To UBound(asTestDataFileNames)
                    sTestDataFileName = asTestDataFileNames(lTestDataIdx)
                    '�����f�[�^
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).asTestLogName(lTestDataIdx) = sTestDataFileName
                    
                    '���҂w�w�w�t�@�C���p�X���X�g
                    If sTestDataFileName <> "-" Then
                        sDicKey = GetFileNameBase(.Name) & "\" & _
                                  gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                  sTestDataFileName
                        sDicItem = lTcShtInfoIdx & "_" & lTcInfoIdx & "_" & lTestDataIdx
                        If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                            'Debug.Print "�d��Key�F" & sDicKey
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

