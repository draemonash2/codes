Attribute VB_Name = "Chk_Summary"
Option Explicit

Private Const SHEET_NAME = "Summary"

Private Enum E_ROW
    ROW_NUM_CHK_A = 4
    ROW_NUM_CHK_B
    ROW_NUM_CHK_C
    ROW_NUM_SUM
End Enum

Private Enum E_CLM
    CLM_ERR = 4
    CLM_WARN
End Enum

Public Enum E_FUNC_TYPE
    FUNC_TYPE_A
    FUNC_TYPE_B
    FUNC_TYPE_C
End Enum

Public Type T_ERROR_MSG_SUMMARY
    lErrNumTestData As Long
    lErrNumExistLogFile As Long
    lErrNumTestDataOmission As Long
    lWarnNumTestData As Long
    lWarnNumExistLogFile As Long
    lWarnNumTestDataOmission As Long
    lErrNumSummary As Long
    lWarnNumSummary As Long
End Type

Private gtErrMsgSum As T_ERROR_MSG_SUMMARY

Public Function ChkSummaryInit()
    Dim tErrMsgSum As T_ERROR_MSG_SUMMARY
    gtErrMsgSum = tErrMsgSum
End Function

Public Function ChkSummaryMain( _
    ByRef wChkRsltBook As Workbook _
)
    Call OutpErrNum(wChkRsltBook)
End Function

Private Function OutpErrNum( _
    ByRef wChkRsltBook As Workbook _
)
    '合計算出
    gtErrMsgSum.lErrNumSummary = _
        gtErrMsgSum.lErrNumTestData + _
        gtErrMsgSum.lErrNumExistLogFile + _
        gtErrMsgSum.lErrNumTestDataOmission
    gtErrMsgSum.lWarnNumSummary = _
        gtErrMsgSum.lWarnNumTestData + _
        gtErrMsgSum.lWarnNumExistLogFile + _
        gtErrMsgSum.lWarnNumTestDataOmission
    
    '+++ シートコピー +++
    Call CopyRsltSht(wChkRsltBook, SHEET_NAME)
    
    '+++ エラー/ワーニング数を出力 +++
    With wChkRsltBook.Sheets(SHEET_NAME)
        .Cells(ROW_NUM_CHK_A, CLM_ERR).Value = gtErrMsgSum.lErrNumTestData
        .Cells(ROW_NUM_CHK_B, CLM_ERR).Value = gtErrMsgSum.lErrNumExistLogFile
        .Cells(ROW_NUM_CHK_C, CLM_ERR).Value = gtErrMsgSum.lErrNumTestDataOmission
        .Cells(ROW_NUM_CHK_A, CLM_WARN).Value = gtErrMsgSum.lWarnNumTestData
        .Cells(ROW_NUM_CHK_B, CLM_WARN).Value = gtErrMsgSum.lWarnNumExistLogFile
        .Cells(ROW_NUM_CHK_C, CLM_WARN).Value = gtErrMsgSum.lWarnNumTestDataOmission
        
        .Cells(ROW_NUM_SUM, CLM_ERR).Value = gtErrMsgSum.lErrNumSummary
        .Cells(ROW_NUM_SUM, CLM_WARN).Value = gtErrMsgSum.lWarnNumSummary
    End With
End Function

Public Function SetErrNum2Summary( _
    ByVal eFuncType As E_FUNC_TYPE, _
    ByVal lErrNum As Long, _
    ByVal lWarnNum As Long _
)
    Select Case eFuncType
        Case FUNC_TYPE_A
            gtErrMsgSum.lErrNumTestData = lErrNum
            gtErrMsgSum.lWarnNumTestData = lWarnNum
        Case FUNC_TYPE_B:
            gtErrMsgSum.lErrNumExistLogFile = lErrNum
            gtErrMsgSum.lWarnNumExistLogFile = lWarnNum
        Case FUNC_TYPE_C:
            gtErrMsgSum.lErrNumTestDataOmission = lErrNum
            gtErrMsgSum.lWarnNumTestDataOmission = lWarnNum
        Case Else
            Stop
    End Select
End Function

Public Function ErrorExist() As Boolean
    ErrorExist = gtErrMsgSum.lErrNumSummary > 0
End Function

Public Function WarningExist() As Boolean
    WarningExist = gtErrMsgSum.lWarnNumSummary > 0
End Function

