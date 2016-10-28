Attribute VB_Name = "ArrayMng"
Option Explicit

' array manage library v1.0

'===========================================================
'= 概要：String 配列に対して Push する。
'= 引数：sPushStr  [in]  String      Push する文字列
'= 引数：asTrgtStr [Out] StrArray    Push 対象配列
'= 戻値：なし
'===========================================================
Public Function PushToStrArray( _
    ByVal sPushStr As String, _
    ByRef asTrgtStr() As String _
)
    Dim lNewArrIdx As Long
 
    If Sgn(asTrgtStr) = 0 Then
        ReDim Preserve asTrgtStr(0)
        asTrgtStr(0) = sPushStr
    Else
        lNewArrIdx = UBound(asTrgtStr) + 1
        ReDim Preserve asTrgtStr(lNewArrIdx)
        asTrgtStr(lNewArrIdx) = sPushStr
    End If
End Function
 
'===========================================================
'= 概要：String 配列に対して Pop する。
'=       初期化なし配列が指定された場合、"" を返却する。
'= 引数：asSrcStr [In]  StrArray    Pop 対象配列
'= 引数：sPopStr  [Out] String      Pop した文字列
'= 戻値：なし
'===========================================================
Public Function PopToStrArray( _
    ByRef asSrcStr() As String, _
    ByRef sPopStr As String _
)
    Dim lPrevArrIdx As Long
 
    If Sgn(asSrcStr) = 0 Then
        sPopStr = ""
    ElseIf UBound(asSrcStr) = 0 Then
        sPopStr = asSrcStr(0)
        ReDim asSrcStr(0)
    Else
        lPrevArrIdx = UBound(asSrcStr)
        sPopStr = asSrcStr(lPrevArrIdx)
        ReDim Preserve asSrcStr(lPrevArrIdx - 1)
    End If
End Function

