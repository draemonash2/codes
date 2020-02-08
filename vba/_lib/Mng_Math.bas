Attribute VB_Name = "Mng_Math"
Option Explicit

' math manage library v1.00

' ==================================================================
' = 概要    余り演算
' =         Mod 演算子は 2,147,483,647 より大きい数字はオーバーフローする。
' =         本関数は上記以上の数値を扱うことができる。
' = 引数    cNum1   Currency    [in]    値1
' = 引数    cNum2   Currency    [in]    値2
' = 戻値            Currency            演算結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Math.bas
' ==================================================================
Public Function ModEx( _
    ByVal cNum1 As Currency, _
    ByVal cNum2 As Currency _
) As Currency
    ModEx = CDec(cNum1) - Fix(CDec(cNum1) / CDec(cNum2)) * CDec(cNum2)
End Function
    Private Sub Test_ModEx()
        Debug.Print "*** test start! ***"
        Debug.Print ModEx(12, 2)             '0
        Debug.Print ModEx(12, 3)             '0
        Debug.Print ModEx(12, 5)             '2
        Debug.Print ModEx(2147483647, 5)     '2
        Debug.Print ModEx(2147483648#, 5)    '3
        Debug.Print ModEx(2147483649#, 5)    '4
        Debug.Print ModEx(-2147483647, 5)    '-2
        Debug.Print ModEx(5, 2147483648#)    '5
        Debug.Print ModEx(5, 2147483649#)    '5
        Debug.Print ModEx(0, 5)              '0
       'Debug.Print ModEx(5, 0)              'プログラム停止
        Debug.Print "*** test finished! ***"
    End Sub

