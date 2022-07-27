Attribute VB_Name = "Mng_Info_TCDoc_4UT"
Option Explicit

Private Const TC_SHT_KEY_WORD = "UT Checklist"
Private Const SRCH_KEYWORD_FILE_NAME = "File Name"
Private Const SRCH_KEYWORD_MODULE_NAME = "Module Name"
Private Const SRCH_KEYWORD_TESTER_SUMMARY = "UT���{�Җ�"
Private Const SRCH_KEYWORD_TC_TESTER = "�@�]���ҁ@"
Private Const SRCH_KEYWORD_TEST_RSLT = "���ʔ���"
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
    
    CLM_PRE_STRT = 1  '�J�n�� - 1
    CLM_SHT_NAME
    CLM_TC_NO
    CLM_WRI_RSLT
    CLM_RESERVE_01 '���g�p
    CLM_PRE_TESTER
    CLM_PRE_TEST_DATE
    CLM_PRE_TEST_RSLT
    CLM_PRE_TEST_DATA
    CLM_PRE_REV
    CLM_RESERVE_02 '���g�p
    CLM_PST_TESTER
    CLM_PST_TEST_DATE
    CLM_PST_TEST_RSLT
    CLM_PST_TEST_DATA
    CLM_PST_REV
    CLM_PST_END '�ŏI�� + 1
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
        '### �����t�F�[�Y�擾 ###
        gtTestDocInfo.eTrgtPhase = TRGT_PHASE_UT
        
        '### �u�b�N���擾 ###
        gtTestDocInfo.sTcDocName = .Name
        
        For lShtIdx = 1 To .Sheets.Count
            '���ڃV�[�g����
            tNearCellData = GetNearCellData(.Sheets(lShtIdx), TC_SHT_KEY_WORD, 0, 0)
            If tNearCellData.bIsCellDataExist = True Then
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
                gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName = .Sheets(lShtIdx).Name
                
                '### �\�[�X�t�@�C�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_FILE_NAME
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sSrcFileName = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                
                '### ���W���[�����擾 ###
                sSrchKeyWord = SRCH_KEYWORD_MODULE_NAME
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sModuleName = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                
                '### UT���{�Җ��擾 ###
                sSrchKeyWord = SRCH_KEYWORD_TESTER_SUMMARY
                tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                If tNearCellData.bIsCellDataExist = True Then
                    gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sTester = tNearCellData.sCellValue
                Else
                    Call StoreErrorMsg( _
                                        "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                        "  �u�b�N���F" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                        "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                        "  �Z���F" & sSrchKeyWord _
                                      )
                End If
                
                '======================
                '=== �Z������擾 ===
                '======================
                tTcRowClmInfo = GetCellClmInfo(wTrgtBook, lShtIdx)
                
                '�G���[�o��
                Call OutpErrorMsg(ERROR_PROC_STOP)
                
                '### ���ڏ��擾 ###
                '���ڐ��O
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
        tTcRowClmInfo.lTcNoClm = tNearCellData.lClm
        tTcRowClmInfo.lTcStrtRow = tNearCellData.lRow + 2
        tTcRowClmInfo.lTcEndRow = .Sheets(lShtIdx).Cells(.Sheets(lShtIdx).Rows.Count, tTcRowClmInfo.lTcNoClm).End(xlUp).Row
        
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
        tTcRowClmInfo.lTesterClm = tNearCellData.lClm
        
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
        tTcRowClmInfo.lTestDateClm = tNearCellData.lClm
        
        '### ���ʔ��� �Z�����擾 ###
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
        tTcRowClmInfo.lTestResultClm = tNearCellData.lClm
        
        '### �����f�[�^ �Z�����擾 ###
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
        tTcRowClmInfo.lTcDataClm = tNearCellData.lClm
        
        '### Rev �Z�����擾 ###
        sSrchKeyWord = SRCH_KEYWORD_REV
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
        '���Ԃ��u-�v���󗓂̏ꍇ�A����
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
            
            '### ���Ԏ擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestCaseNo = sTcNo
            
            '### �]���Ҏ擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTester = sTester
            
            '### �N�����擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestDate = sTestDate
            
            '### ���ʔ���擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestResult = sTestResult
            
            '### Rev�擾 ###
            gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).atTcInfo(lTcInfoIdx).sTestRevSrc = sRev
            
            '### �����f�[�^�{���҃t�@�C���p�X���X�g�擾 ###
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
                        Select Case GetFileNameExt(sTestDataFileName)
                            Case "csv"
                                sDicKey = GetFileNameBase(.Name) & "\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                          sTestDataFileName
                                sDicItem = lTcShtInfoIdx & "_" & lTcInfoIdx & "_" & lTestDataIdx
                                If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                                    'Debug.Print "�d��Key�F" & sDicKey
                                Else
                                    gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                                End If
                            Case "htm"
                                sDicKey = GetFileNameBase(.Name) & "\" & _
                                          gtTestDocInfo.atTcShtInfo(lTcShtInfoIdx).sShtName & "\" & _
                                          sTestDataFileName
                                sDicItem = lTcShtInfoIdx
                                If gtTestDocInfo.oLogExpPathList.exists(sDicKey) = True Then
                                    'Debug.Print "�d��Key�F" & sDicKey
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
                                    'Debug.Print "�d��Key�F" & sDicKey
                                Else
                                    gtTestDocInfo.oLogExpPathList.Add sDicKey, sDicItem
                                End If
                            Case Else
                                '�����ł̓`�F�b�N���Ȃ��B
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
    Dim oTestedTcList As Object 'Key:[�V�[�g��]_[����]�AItem:���g�p�iTrue�Œ�j
    
    '=== �������{�ςݍ��ڃ��X�g�쐬 ===
    Set oTestedTcList = CreTestedTcList
    
    '=== �������ʗ��������� ===
    Call WriTestRslt(wTcDocBook, oTestedTcList)
    
    '=== �������ʗ��������݌��ʏo�� ===
    Call OutpWriRslt(wWriRsltBook)
End Function

Private Function CreTestedTcList() As Object
    Dim sPath As String
    Dim sRelativeFilePath As String
    Dim eSvnModStatus As E_SVN_MOD_STATUS
    Dim sFileBaseName As String
    Dim atSvnModStatInfo() As T_SVN_MOD_STAT_INFO
    Dim lSvnModStatInfoIdx As Long
    Dim oTestedTcList As Object 'Key:[�V�[�g��]_[����]�AItem:���g�p�iTrue�Œ�j
    
    Set oTestedTcList = CreateObject("Scripting.Dictionary")
    
    'SVN �̕ύX��ԃ��X�g�擾
    atSvnModStatInfo = GetSvnModStatList(gtInputInfo.sTestLogDirPath)
    
    For lSvnModStatInfoIdx = 0 To UBound(atSvnModStatInfo)
    '�ύX�ςݏ�ԃ��X�g
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
                    '�������{�ς݃��X�g�ǉ�
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
    
    '*** �������݌��ʁi�������ڏ����A���O�t�H���_�p�X�j�i�[ ***
    gtWriRsltInfo.sTcDocFileName = GetFileName(gtInputInfo.sTestDocFilePath)
    gtWriRsltInfo.sLogDirPath = gtInputInfo.sTestLogDirPath
    
    With wTcDocBook
        For lShtIdx = 1 To .Sheets.Count
            '���ڃV�[�g����
            tNearCellData = GetNearCellData(.Sheets(lShtIdx), TC_SHT_KEY_WORD, 0, 0)
            If tNearCellData.bIsCellDataExist = True Then
                '�e���ڂ̗�ԍ��擾
                tTcRowClmInfo = GetCellClmInfo(wTcDocBook, lShtIdx)
                
                '### ���ڏ��擾 ###
                '���ڐ��O
                lTcNum = tTcRowClmInfo.lTcEndRow - tTcRowClmInfo.lTcStrtRow + 1
                If lTcNum = 0 Then
                    'Do Nothing
                Else
                    sSheetName = .Sheets(lShtIdx).Name
                    
                    '���W���[����
                    sSrchKeyWord = SRCH_KEYWORD_MODULE_NAME
                    tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                    If tNearCellData.bIsCellDataExist = True Then
                        sModuleName = tNearCellData.sCellValue
                    Else
                        Call StoreErrorMsg( _
                                            "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                            "  �u�b�N���F" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                            "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                            "  �Z���F" & sSrchKeyWord _
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
                        '�����ς݂�����
                        If oTestedTcList.exists(sSheetName & "_" & sTestCaseNo) = True Then
                            bIsTested = True
                            sTester = gtInputInfo.sTester
                            sTestDate = gtInputInfo.sTestDate
                            sTestRslt = gtInputInfo.sTestRslt
                            sTcData = _
                                sSheetName & "_" & sTestCaseNo & ".csv" & vbLf & _
                                sModuleName & ".txt" & vbLf & _
                                "�e�X�g���ʕ񍐏�.htm"
                            sRevSrc = gtInputInfo.sRevSrc
                            
                            '*** �������݌��ʊi�[ ***
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sSheetName = sSheetName
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sTestCaseNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm).Value
                            If .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value = "" Then
                                gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "�V�K"
                            Else
                                gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "�㏑��"
                            End If
                            '�������ݑO
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            '�������݌�
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTester = sTester
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestDate = sTestDate
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt = sTestRslt
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestData = sTcData
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostRevSrc = sRevSrc
                            
                            '*** �����f�[�^���������� ***
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value = sTester
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value = sTestDate
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value = sTestRslt
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value = sTcData
                            .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value = sRevSrc
                        Else
                            '*** �������݌��ʊi�[ ***
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sSheetName = sSheetName
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sTestCaseNo = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcNoClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sWriRslt = "�ύX�Ȃ�"
                            '�������ݑO
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPreRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            '�������݌�
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTester = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTesterClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestDate = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestDateClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestRslt = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestResultClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostTestData = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTcDataClm).Value
                            gtWriRsltInfo.atWriRsltInfoRow(lWriRsltInfoRowIdx).sPostRevSrc = .Sheets(lShtIdx).Cells(lRowIdx, tTcRowClmInfo.lTestRevClm).Value
                            
                            '*** �����f�[�^���������� ***
                            '�����ς݂łȂ����߁A���������Ȃ�
                        End If
                    Next lRowIdx
                    
                    '�ꍀ�ڂł������ς݂ł����
                    If bIsTested = True Then
                        'UT���{�Җ�
                        sSrchKeyWord = SRCH_KEYWORD_TESTER_SUMMARY
                        tNearCellData = GetNearCellData(.Sheets(lShtIdx), sSrchKeyWord, 0, 1)
                        If tNearCellData.bIsCellDataExist = True Then
                            .Sheets(lShtIdx).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = gtInputInfo.sTester
                        Else
                            Call StoreErrorMsg( _
                                                "�ȉ��̃Z����������܂���ł����I" & vbNewLine & _
                                                "  �u�b�N���F" & gtTestDocInfo.sTcDocName & vbNewLine & _
                                                "  �V�[�g���F" & .Sheets(lShtIdx).Name & vbNewLine & _
                                                "  �Z���F" & sSrchKeyWord _
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

'���������e�X�g�I������
Private Function OutpWriRslt( _
    ByRef wWriRsltBook As Workbook _
)
    Dim lWriRsltInfoRowIdx As Long
    Dim lRowIdx As Long
    
    '+++ �V�[�g�R�s�[ +++
    Call CopyRsltSht(wWriRsltBook, RSLT_SHEET_NAME)
    
    '+++ �������݌��ʏo�� +++
    With wWriRsltBook.Sheets(RSLT_SHEET_NAME)
        '### �������ڏ��t�@�C���� ###
        .Cells(ROW_TC_DOC_NAME, CLM_TC_DOC_NAME).Value = gtWriRsltInfo.sTcDocFileName
        
        '### �������O�t�H���_�� ###
        .Cells(ROW_LOG_DIR_PATH, CLM_LOG_DIR_PATH).Value = gtWriRsltInfo.sLogDirPath
        
        '### �s��� ###
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

'�������e�X�g�p������
Sub test()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\test\11_TF2���v���g2nd\31_�P�̎������O\02_�f�f���j�^�ύX"
    Call OutpTestResult4UT(Nothing, Nothing)
End Sub
