Attribute VB_Name = "Mng_Info_TCDoc_Cmn"
Option Explicit

Public Const SRCH_KEYWORD_TC_NO = "����"
Public Const SRCH_KEYWORD_TEST_DATE = "�N����"
Public Const SRCH_KEYWORD_TEST_DATA = "�����f�[�^"

'===============================
'= �������ڏ��f�[�^���͗p�\����
'===============================
Private Type T_TESTCASE_INFO
    sTestCaseNo As String
    asTestLogName() As String
    sTestDataCellValue As String
    sTester As String
    sTestDate As String
    sTestResult As String
    sTestRevSrc As String    '�u�P�̎����v�I�����̂ݎg�p
    sTestRevHexAbs As String '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
    sTestRevA2L As String    '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
End Type

Private Type T_SHEET_INFO
    sShtName As String
    sSrcFileName As String    '�u�P�̎����v�I�����̂ݎg�p
    sModuleName As String     '�u�P�̎����v�I�����̂ݎg�p
    sTester As String         '�u�P�̎����v�I�����̂ݎg�p
    atTcInfo() As T_TESTCASE_INFO
End Type

Private Type T_TEST_DOC_INFO
    eTrgtPhase As E_TRGT_PHASE
    sTcDocName As String
    atTcShtInfo() As T_SHEET_INFO
    oLogExpPathList As Object 'Key:���҃t�@�C���p�X Item:�����f�[�^�L�ڗL��
End Type

'===============================
'= �������ڏ��������݌��ʊi�[�p�\����
'===============================
Private Type T_WRI_RSLT_INFO_ROW
    sSheetName As String
    sTestCaseNo As String
    sWriRslt As String
    sPreTester As String
    sPreTestDate As String
    sPreTestRslt As String
    sPreTestData As String
    sPreRevSrc As String     '�u�P�̎����v�I�����̂ݎg�p
    sPreRevHexAbs As String  '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
    sPreRevA2L As String     '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
    sPostTester As String
    sPostTestDate As String
    sPostTestRslt As String
    sPostTestData As String
    sPostRevSrc As String    '�u�P�̎����v�I�����̂ݎg�p
    sPostRevHexAbs As String '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
    sPostRevA2L As String    '�u�����E�@�\�E�V�X�e�������v�I�����̂ݎg�p
End Type

Private Type T_WRI_RSLT_INFO
    sTcDocFileName As String
    sLogDirPath As String
    atWriRsltInfoRow() As T_WRI_RSLT_INFO_ROW
End Type

Public gtWriRsltInfo As T_WRI_RSLT_INFO
Public gtTestDocInfo As T_TEST_DOC_INFO

Public Function TcDocInfoInit()
    Dim tTestDocInfo As T_TEST_DOC_INFO
    Dim tWriRsltInfo As T_WRI_RSLT_INFO
    gtTestDocInfo = tTestDocInfo
    gtWriRsltInfo = tWriRsltInfo
    Set gtTestDocInfo.oLogExpPathList = CreateObject("Scripting.Dictionary")
End Function

Public Function GetTCDocInfo()
    Dim wTrgtBook As Workbook
    
    '���ڏ��I�[�v��
    Set wTrgtBook = ExcelFileOpen(gtInputInfo.sTestDocFilePath)
    
    Select Case gtInputInfo.eTrgtPhase
        Case TRGT_PHASE_UT: Call GetTCDocInfo4UT(wTrgtBook)
        Case TRGT_PHASE_CT: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case TRGT_PHASE_FT: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case TRGT_PHASE_ST: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case Else:          Stop
    End Select
    
    '���ڏ��N���[�Y
    Call ExcelFileClose(wTrgtBook, False)
End Function

Public Function WriTestRslt( _
    ByRef wTcDocBook As Workbook, _
    ByRef wWriRsltBook As Workbook _
)
    Select Case gtInputInfo.eTrgtPhase
        Case TRGT_PHASE_UT: Call OutpTestResult4UT(wTcDocBook, wWriRsltBook)
        'Case TRGT_PHASE_CT: Call OutpTestResult4CTFTST(wTcDocBook, wWriRsltBook)
        'Case TRGT_PHASE_FT: Call OutpTestResult4CTFTST(wTcDocBook, wWriRsltBook)
        'Case TRGT_PHASE_ST: Call OutpTestResult4CTFTST(wTcDocBook, wWriRsltBook)
        Case Else:          Stop
    End Select
End Function

Public Function TcDocInfoTerminate()
    Set gtTestDocInfo.oLogExpPathList = Nothing
End Function

Public Function GetTestDataArray( _
    ByVal sTrgtStr As String _
) As String()
    Dim asRetArray() As String
    
    '���s CR �͎��O�ɍ폜
    sTrgtStr = Replace(sTrgtStr, vbCr, "")
    '�A���������s LF �͈�ɂ܂Ƃ߂�
    Do While InStr(sTrgtStr, vbLf & vbLf)
        sTrgtStr = Replace(sTrgtStr, vbLf & vbLf, vbLf)
    Loop
    '�����̉��s LF ���폜
    If Right(sTrgtStr, 1) = vbLf Then
        sTrgtStr = Left(sTrgtStr, Len(sTrgtStr) - 1)
    Else
        'Do Nothing
    End If
    '�擪�̉��s LF ���폜
    If Left(sTrgtStr, 1) = vbLf Then
        sTrgtStr = Right(sTrgtStr, Len(sTrgtStr) - 1)
    Else
        'Do Nothing
    End If
    
    If sTrgtStr = "" Or sTrgtStr = "-" Then
        ReDim Preserve asRetArray(0)
        asRetArray(0) = sTrgtStr
    Else
        '���s LF �ŕ���
        asRetArray = Split(sTrgtStr, vbLf)
    End If
    
    GetTestDataArray = asRetArray
End Function

'���e�X�g�p��
Sub test()
    Dim sTrgtStr As String
    Dim asTrgtStr() As String
'    sTrgtStr = _
'                "" & vbCrLf & _
'                "" & vbCrLf & _
'                "" & vbCrLf & _
'                "aaaa" & vbLf & _
'                "bbb" & vbLf & _
'                "ccc" & vbCrLf & _
'                "" & vbCrLf & _
'                "d" & vbCrLf
    sTrgtStr = "-"
    asTrgtStr = GetTestDataArray(sTrgtStr)
End Sub


