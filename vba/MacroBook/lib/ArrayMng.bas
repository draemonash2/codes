Attribute VB_Name = "Mng_Array"
Option Explicit

' array manage library v1.1

Public Enum E_INSERT_TYPE
    E_INSERT_TOP
    E_INSERT_MIDDLE
    E_INSERT_BOTTOM
End Enum

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

'===========================================================
'= 概要：String 配列に対して指定位置に配列を挿入する。
'= 引数：eInsertType    [In]        Enum        挿入種別（先頭/中間/末尾）
'= 引数：lTrgtIdx       [In]        Long        挿入したい要素番号
'= 引数：asTrgtStr()    [Out]       String      挿入したい文字配列
'= 引数：asBaseStr()    [In,Out]    String()    挿入元文字配列、挿入後の文字配列
'= 戻値：なし
'= 覚書：
'=         例１）配列番号 0～3 の配列に対して、挿入種別「先頭」を
'=               指定した場合
'=                 0, 1, 2, 3 ⇒ _, 0, 1, 2, 3
'=         例２）配列番号 0～3 の配列に対して、挿入種別「末尾」を
'=               指定した場合
'=                 0, 1, 2, 3 ⇒ 0, 1, 2, 3, _
'=         例３）配列番号 0～3 の配列に対して、挿入種別「中間」、
'=               lTrgtIdx = 2 を指定した場合
'=                 0, 1, 2, 3 ⇒ 0, 1, _, 2, 3
'===========================================================
Public Function InsertToStrArray( _
    ByRef eInsertType As E_INSERT_TYPE, _
    ByRef lTrgtIdx As Long, _
    ByRef asTrgtStr() As String, _
    ByRef asBaseStr() As String _
)

End Function

'===========================================================
'= 概要：String 配列に対して指定位置に配列を挿入する。（指定要素置き換え）
'= 引数：lTrgtIdx       [In]        Long        置き換えたい要素番号
'= 引数：asRepArr()     [In]        String()    置き換えたい文字配列
'= 引数：asBaseArr()    [In,Out]    String()    置き換え元文字配列、挿入後の文字配列
'= 戻値：                           Boolean     置き換え結果
'= 覚書：
'=         例１）配列A（0～4）に対して、配列B（要素0～2)、
'=               lTrgtIdx = 2 を指定した場合
'=                     0     1     2     3     4
'=                 A = A[0], A[1], A[2], A[3], A[4]
'=                 ↓
'=                     0     1     2     3     4     5     6
'=                 A = A[0], A[1], B[0], B[1], B[2], A[3], A[4]
'===========================================================
Public Function InsRepToStrArray( _
    ByRef lTrgtIdx As Long, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String _
) As Boolean
    Dim bIsError As Boolean
    bIsError = False
    
    '引数チェック
    Debug.Assert Sgn(asRepArr) <> 0
    Debug.Assert Sgn(asBaseArr) <> 0
    Debug.Assert lTrgtIdx >= LBound(asBaseArr) And lTrgtIdx <= UBound(asBaseArr)
    Debug.Assert LBound(asRepArr) = 0
    Debug.Assert LBound(asBaseArr) = 0
    
    ReDim Preserve asBaseArr(UBound(asBaseArr) + UBound(asRepArr))
    '移動
    Dim lBaseSrcIdx As Long
    Dim lBaseDstIdx As Long
    For lBaseDstIdx = UBound(asBaseArr) To (lTrgtIdx + UBound(asRepArr) + 1) Step -1
        lBaseSrcIdx = lBaseDstIdx - UBound(asRepArr)
        asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
    Next lBaseDstIdx
    
    '挿入
    Dim lBaseIdx As Long
    Dim lRepIdx As Long
    For lBaseIdx = lTrgtIdx To (lTrgtIdx + UBound(asRepArr))
        asBaseArr(lBaseIdx) = asRepArr(lRepIdx)
        lRepIdx = lRepIdx + 1
    Next lBaseIdx
End Function
    Private Function Test_InsRepToStrArray()
        Dim asBaseArr() As String
        Dim asRepArr() As String
        
        'asBaseArr(3),asRepArr(2)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(2, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(3, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(4, asRepArr, asBaseArr)
        
        'asBaseArr(3),asRepArr(0)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(2, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(3, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(4, asRepArr, asBaseArr)

        'asBaseArr(0),asRepArr(3)
        Call Test_InsRepToStrArraySub03(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub03(asBaseArr, asRepArr): Call InsRepToStrArray(1, asRepArr, asBaseArr)

        'asBaseArr(0),asRepArr(0)
        Call Test_InsRepToStrArraySub04(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub04(asBaseArr, asRepArr): Call InsRepToStrArray(1, asRepArr, asBaseArr)

    End Function
        Private Function Test_InsRepToStrArraySub01( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(3)
            ReDim Preserve asRepArr(2)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            asBaseArr(3) = "3"
            asRepArr(0) = "a"
            asRepArr(1) = "b"
            asRepArr(2) = "c"
        End Function
        Private Function Test_InsRepToStrArraySub02( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(3)
            ReDim Preserve asRepArr(0)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            asBaseArr(3) = "3"
            asRepArr(0) = "a"
        End Function
        Private Function Test_InsRepToStrArraySub03( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(0)
            ReDim Preserve asRepArr(3)
            asBaseArr(0) = "0"
            asRepArr(0) = "a"
            asRepArr(1) = "b"
            asRepArr(2) = "c"
            asRepArr(3) = "d"
        End Function
        Private Function Test_InsRepToStrArraySub04( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(0)
            ReDim Preserve asRepArr(0)
            asBaseArr(0) = "0"
            asRepArr(0) = "a"
        End Function
