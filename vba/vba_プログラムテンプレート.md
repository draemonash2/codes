```vba
Option Explicit

Private Const SHEET_NAME = "XXX"
Private Const REF_CELL_KEYWORD = "YYY"

Private Enum E_ROW_OFFSET
    ROW_AAA = 0
    ROW_BBB
End Enum

Private Enum E_CLM_OFFSET
    CLM_CCC = 0
    CLM_DDD
End Enum

Private Type T_XXX
    lXxxStrtRow As Long
    lXxxStrtClm As Long
End Type

Public gtSampleStruct As T_XXX

Public Function SampleFuncInit()
    Dim tSampleStructInit As T_XXX
    gtSampleStruct = tSampleStructInit
End Function

Public Function SampleFuncTerminate()
    'ÅöèIóπèàóù
End Function

Public Function SampleFunc()
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ThisWorkbook.Sheets(SHEET_NAME)
    
    Dim rFindResult As Range
    Dim sSrchKeyword As String
    Dim lSrchCellRow As Long
    Dim lSrchCellClm As Long
    sSrchKeyword = REF_CELL_KEYWORD
    Set rFindResult = shTrgtSht.Cells.Find(sSrchKeyword, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        MsgBox sSrchKeyword & "Ç™å©Ç¬Ç©ÇËÇ‹ÇπÇÒÇ≈ÇµÇΩ"
    Else
        lSrchCellRow = rFindResult.Row
        lSrchCellClm = rFindResult.Column
    End If
End Function
```
