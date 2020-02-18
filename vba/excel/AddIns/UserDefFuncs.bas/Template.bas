Attribute VB_Name = "Template"
Option Explicit

' templates v1.0

' =============================================================================
' = 概要    ★
' = 覚書    なし
' = 依存    XXX.bas/Xxxx()
' =         YYY.bas/Yyyy()
' = 所属    ★
' =============================================================================
Private Sub TemplateSub()
    '▼▼▼設定 ここから▼▼▼
    Const sMACRO_NAME = "★マクロ名★"
    '▲▲▲設定 ここまで▲▲▲
    
    Dim vCalcSetting As Variant
    Application.ScreenUpdating = False
    vCalcSetting = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim rFindResult As Range
    Dim sSrchKeyword As String
    
    With ThisWorkbook.Sheets("★シート名★")
        '*** 事前処理 ***
        '開始終了行,列検索
        sSrchKeyword = "★検索単語★"
        Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "が見つかりませんでした", vbCritical, sMACRO_NAME
            MsgBox "処理を中断します", vbCritical, sMACRO_NAME
            Exit Sub
        End If
        Dim lTitleRow As Long
        Dim lStrtRow As Long
        Dim lLastRow As Long
        Dim lStrtClm As Long
        Dim lLastClm As Long
        lTitleRow = rFindResult.Row
        lStrtRow = rFindResult.Row + 1
        lStrtClm = rFindResult.Column
        lLastRow = .Cells(.Rows.Count, lStrtClm).End(xlUp).Row
        lLastClm = .Cells(lTitleRow, .Columns.Count).End(xlToLeft).Column
        
        '*** 本処理 ***
        Dim lRowIdx As Long
        For lRowIdx = lStrtRow To lLastRow
            '★ここに処理を書く★
        Next lRowIdx
    End With
  
    Application.Calculation = vCalcSetting
    Application.ScreenUpdating = True
End Sub

' ==================================================================
' = 概要    ★
' = 引数    sAAAA           String   [in]   ★入力文字列★
' = 戻値                    Boolean         ★戻り値★
' = 覚書    なし
' = 依存    なし
' = 所属    XXX.bas
' ==================================================================
Private Function TemplateFunc()
    '★ここに処理を書く★
End Function
    Private Sub Test_TemplateFunc()
        Dim bRet As Boolean
        Debug.Print "*** test start! ***"
        '★ここに処理を書く★
        Debug.Print "*** test finished! ***"
    End Sub


