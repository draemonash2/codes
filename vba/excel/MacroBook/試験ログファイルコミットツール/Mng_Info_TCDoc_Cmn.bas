Attribute VB_Name = "Mng_Info_TCDoc_Cmn"
Option Explicit

Public Const SRCH_KEYWORD_TC_NO = "項番"
Public Const SRCH_KEYWORD_TEST_DATE = "年月日"
Public Const SRCH_KEYWORD_TEST_DATA = "試験データ"

'===============================
'= 試験項目書データ入力用構造体
'===============================
Private Type T_TESTCASE_INFO
    sTestCaseNo As String
    asTestLogName() As String
    sTestDataCellValue As String
    sTester As String
    sTestDate As String
    sTestResult As String
    sTestRevSrc As String    '「単体試験」選択時のみ使用
    sTestRevHexAbs As String '「結合・機能・システム試験」選択時のみ使用
    sTestRevA2L As String    '「結合・機能・システム試験」選択時のみ使用
End Type

Private Type T_SHEET_INFO
    sShtName As String
    sSrcFileName As String    '「単体試験」選択時のみ使用
    sModuleName As String     '「単体試験」選択時のみ使用
    sTester As String         '「単体試験」選択時のみ使用
    atTcInfo() As T_TESTCASE_INFO
End Type

Private Type T_TEST_DOC_INFO
    eTrgtPhase As E_TRGT_PHASE
    sTcDocName As String
    atTcShtInfo() As T_SHEET_INFO
    oLogExpPathList As Object 'Key:期待ファイルパス Item:試験データ記載有無
End Type

'===============================
'= 試験項目書書き込み結果格納用構造体
'===============================
Private Type T_WRI_RSLT_INFO_ROW
    sSheetName As String
    sTestCaseNo As String
    sWriRslt As String
    sPreTester As String
    sPreTestDate As String
    sPreTestRslt As String
    sPreTestData As String
    sPreRevSrc As String     '「単体試験」選択時のみ使用
    sPreRevHexAbs As String  '「結合・機能・システム試験」選択時のみ使用
    sPreRevA2L As String     '「結合・機能・システム試験」選択時のみ使用
    sPostTester As String
    sPostTestDate As String
    sPostTestRslt As String
    sPostTestData As String
    sPostRevSrc As String    '「単体試験」選択時のみ使用
    sPostRevHexAbs As String '「結合・機能・システム試験」選択時のみ使用
    sPostRevA2L As String    '「結合・機能・システム試験」選択時のみ使用
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
    
    '項目書オープン
    Set wTrgtBook = ExcelFileOpen(gtInputInfo.sTestDocFilePath)
    
    Select Case gtInputInfo.eTrgtPhase
        Case TRGT_PHASE_UT: Call GetTCDocInfo4UT(wTrgtBook)
        Case TRGT_PHASE_CT: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case TRGT_PHASE_FT: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case TRGT_PHASE_ST: Call GetTCDocInfo4CTFTST(wTrgtBook)
        Case Else:          Stop
    End Select
    
    '項目書クローズ
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
    
    '改行 CR は事前に削除
    sTrgtStr = Replace(sTrgtStr, vbCr, "")
    '連続した改行 LF は一つにまとめる
    Do While InStr(sTrgtStr, vbLf & vbLf)
        sTrgtStr = Replace(sTrgtStr, vbLf & vbLf, vbLf)
    Loop
    '末尾の改行 LF を削除
    If Right(sTrgtStr, 1) = vbLf Then
        sTrgtStr = Left(sTrgtStr, Len(sTrgtStr) - 1)
    Else
        'Do Nothing
    End If
    '先頭の改行 LF を削除
    If Left(sTrgtStr, 1) = vbLf Then
        sTrgtStr = Right(sTrgtStr, Len(sTrgtStr) - 1)
    Else
        'Do Nothing
    End If
    
    If sTrgtStr = "" Or sTrgtStr = "-" Then
        ReDim Preserve asRetArray(0)
        asRetArray(0) = sTrgtStr
    Else
        '改行 LF で分割
        asRetArray = Split(sTrgtStr, vbLf)
    End If
    
    GetTestDataArray = asRetArray
End Function

'★テスト用★
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


