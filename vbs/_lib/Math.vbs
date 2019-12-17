Option Explicit

' ==================================================================
' = 概要    切り上げを行う
' = 引数    dValue      Double  [in]  入力値
' = 引数    lDigit      Long    [in]  桁
' = 戻値                Long          切り上げ結果
' = 覚書    ・負値は未対応
' = 依存    なし
' = 所属    Math.vbs
' ==================================================================
Public Function RoundUp( _
    ByVal dValue, _
    ByVal lDigit _
)
    RoundUp = Fix((dValue + (9 * (10 ^ (-1 * (lDigit + 1))))) * (10 ^ lDigit)) / (10 ^ lDigit)
End Function
    Call Test_RoundUp()
    Private Sub Test_RoundUp()
        Dim Result
        Dim lPaddingLen
        Result = "[Result]"
        Result = Result & vbNewLine & RoundUp(1, 0)    ' 1
        Result = Result & vbNewLine & RoundUp(1.01, 0) ' 1
        Result = Result & vbNewLine & RoundUp(1.1, 0)  ' 2
        Result = Result & vbNewLine & RoundUp(1.5, 0)  ' 2
        Result = Result & vbNewLine & RoundUp(1.6, 0)  ' 2
        Result = Result & vbNewLine & RoundUp(1.9, 0)  ' 2
        Result = Result & vbNewLine & RoundUp(1.99, 0) ' 2
        Result = Result & vbNewLine & RoundUp(2.0, 0)  ' 2
        Result = Result & vbNewLine & RoundUp(2.1, 0)  ' 3
        Result = Result & vbNewLine & RoundUp(0, 0)    ' 0
        Result = Result & vbNewLine & RoundUp(0.1, 0)  ' 1
        Result = Result & vbNewLine & RoundUp(-0.1, 0) ' 未対応
        Result = Result & vbNewLine & RoundUp(-1.5, 0) ' 未対応
        MsgBox Result
    End Sub

