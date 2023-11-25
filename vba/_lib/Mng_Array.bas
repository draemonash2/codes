Attribute VB_Name = "Mng_Array"
Option Explicit

' array manage library v1.34

Public Enum E_INSERT_TYPE
    E_INSERT_TOP
    E_INSERT_MIDDLE
    E_INSERT_BOTTOM
End Enum

' ==================================================================
' = 概要    String 配列に対して Push する。
' = 引数    sPushStr  [in]  String      Push する文字列
' = 引数    asTrgtStr [Out] StrArray    Push 対象配列
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
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
 
' ==================================================================
' = 概要    String 配列に対して Pop する。
' =         初期化なし配列が指定された場合、"" を返却する。
' = 引数    asSrcStr [In]  StrArray    Pop 対象配列
' = 引数    sPopStr  [Out] String      Pop した文字列
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
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

' ==================================================================
' = 概要    String 配列に対して指定位置に配列を挿入する。
' = 引数    eInsertType    [In]        Enum        挿入種別（先頭/中間/末尾）
' = 引数    lTrgtIdx       [In]        Long        挿入したい要素番号
' = 引数    asTrgtStr()    [Out]       String      挿入したい文字配列
' = 引数    asBaseStr()    [In,Out]    String()    挿入元文字配列、挿入後の文字配列
' = 戻値    なし
' = 覚書    
' =         例１）配列番号 0～3 の配列に対して、挿入種別「先頭」を
' =               指定した場合
' =                 0, 1, 2, 3 ⇒ _, 0, 1, 2, 3
' =         例２）配列番号 0～3 の配列に対して、挿入種別「末尾」を
' =               指定した場合
' =                 0, 1, 2, 3 ⇒ 0, 1, 2, 3, _
' =         例３）配列番号 0～3 の配列に対して、挿入種別「中間」、
' =               lTrgtIdx = 2 を指定した場合
' =                 0, 1, 2, 3 ⇒ 0, 1, _, 2, 3
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
Public Function InsertToStrArray( _
    ByRef eInsertType As E_INSERT_TYPE, _
    ByRef lTrgtIdx As Long, _
    ByRef asTrgtStr() As String, _
    ByRef asBaseStr() As String _
)
    'TODO：要実装
End Function

' ==================================================================
' = 概要    String 配列に対して指定位置に配列を挿入する。（指定要素置き換え）
' = 引数    lTrgtIdx       [In]        Long        置き換えたい要素番号
' = 引数    asRepArr()     [In]        String()    置き換えたい文字配列
' = 引数    asBaseArr()    [In,Out]    String()    置き換え元文字配列、挿入後の文字配列
' = 戻値                               Boolean     置き換え結果
' = 覚書    
' =           例１）配列A（asBaseArr要素0～4）に対して、配列B（asRepArr要素0～2)、
' =                 lTrgtIdx = 2 を指定した場合
' =                       0     1     2     3     4
' =                   A = A[0], A[1], A[2], A[3], A[4]
' =                   ↓
' =                       0     1     2     3     4     5     6
' =                   A = A[0], A[1], B[0], B[1], B[2], A[3], A[4]
' =           例２）配列A（asBaseArr要素0～4）に対して、配列B（asRepArr空配列)、
' =                 lTrgtIdx = 2 を指定した場合
' =                       0     1     2     3     4
' =                   A = A[0], A[1], A[2], A[3], A[4]
' =                   ↓
' =                       0     1     2     3
' =                   A = A[0], A[1], A[3], A[4]
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
Public Function InsRepToStrArray( _
    ByRef lTrgtIdx As Long, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String _
) As Boolean
    Dim bIsError As Boolean
    bIsError = False
    
    '引数チェック
    'Debug.Assert Sgn(asRepArr) <> 0
    'Debug.Assert LBound(asRepArr) = 0
    Debug.Assert Sgn(asBaseArr) <> 0
    Debug.Assert LBound(asBaseArr) = 0
    Debug.Assert lTrgtIdx >= LBound(asBaseArr) And lTrgtIdx <= UBound(asBaseArr)
    
    Dim lBaseSrcIdx As Long
    Dim lBaseDstIdx As Long
    
    'asRepArr が未初期化配列の場合、要素番号 lTrgtIdx を削除する
    If Sgn(asRepArr) = 0 Then
        For lBaseSrcIdx = (lTrgtIdx + 1) To UBound(asBaseArr)
            lBaseDstIdx = lBaseSrcIdx - 1
            asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
        Next lBaseSrcIdx
        ReDim Preserve asBaseArr(UBound(asBaseArr) - 1)
    
    'asRepArr が初期化済み配列の場合、要素番号 lTrgtIdx に asRepArr を挿入する
    Else
        ReDim Preserve asBaseArr(UBound(asBaseArr) + UBound(asRepArr))
        '移動
        For lBaseDstIdx = UBound(asBaseArr) To (lTrgtIdx + UBound(asRepArr) + 1) Step -1
            lBaseSrcIdx = lBaseDstIdx - UBound(asRepArr)
            asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
        Next lBaseDstIdx
        
        '挿入
        Dim lBaseIdx As Long
        Dim lRepIdx As Long
        For lBaseIdx = lTrgtIdx To (lTrgtIdx + UBound(asRepArr))
            'asRepArrの要素が空文字列の場合、asBaseArr(lBaseIdx) = asRepArr(lRepIdx) とすると
            'テキストファイル出力時に（なぜか）エラーが発生する。
            'そのため、空文字判定を実施。
            If asRepArr(lRepIdx) = "" Then
                asBaseArr(lBaseIdx) = ""
            Else
                asBaseArr(lBaseIdx) = asRepArr(lRepIdx)
            End If
            lRepIdx = lRepIdx + 1
        Next lBaseIdx
    End If
End Function
    Private Function Test_InsRepToStrArray()
        Dim asBaseArr() As String
        Dim asRepArr() As String
        Dim asRepArr05() As String
        
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
        
        'asBaseArr(2),asRepArr(未初期化)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(0, asRepArr05, asBaseArr)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(1, asRepArr05, asBaseArr)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(2, asRepArr05, asBaseArr)
        
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
        Private Function Test_InsRepToStrArraySub05( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(2)
            'ReDim asRepArr(0)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            'asRepArr(0) = "a"
        End Function

' ==================================================================
' = 概要    挿入先配列の中身からキーワードを検索して挿入配列で置換する
' =         （テンプレートファイルを元にファイルを生成する際に使用する）
' = 引数    sKeyword        String      [in]    キーワード
' =         asRepArr()      String()    [in]    挿入配列
' =         asBaseArr()     String()    [out]   挿入先配列
' =         bIsWholeMatch   Boolean     [in]    キーワード完全一致/部分一致（True:完全一致）
' = 戻値                    Boolean             一致結果
' = 覚書    asBaseArr の中に同じ  が複数含まれている場合、
' =         先頭の sKeyword のみ置き換える
' = 依存    Mng_Array.bas/InsRepToStrArray()
' = 所属    Mng_Array.bas
' ==================================================================
Public Function ReplaceArray( _
    ByVal sKeyword As String, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String, _
    Optional ByVal bIsWholeMatch As Boolean = True _
) As Boolean
    Dim lIdx As Long
    Dim bIsMatch As Boolean
    bIsMatch = False
    '完全一致
    If bIsWholeMatch = True Then
        For lIdx = LBound(asBaseArr) To UBound(asBaseArr)
            If asBaseArr(lIdx) = sKeyword Then
                Call InsRepToStrArray(lIdx, asRepArr, asBaseArr)
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lIdx
    '部分一致
    Else
        For lIdx = LBound(asBaseArr) To UBound(asBaseArr)
            If InStr(asBaseArr(lIdx), sKeyword) Then
                Call InsRepToStrArray(lIdx, asRepArr, asBaseArr)
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lIdx
    End If
    ReplaceArray = bIsMatch
End Function
    Private Sub Test_ReplaceArray()
        Dim asBaseFileLine() As String
        Dim asRepLine01() As String
        Dim asRepLine02() As String
        Dim asRepLine03() As String
        Dim asRepLine04() As String
        Dim asRepLine05() As String
        Dim sKeyword As String
        Dim bRet As Boolean
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
         'デスクトップフォルダ
        asBaseFileLine = InputTxtFile(objWshShell.SpecialFolders("Desktop") & "\" & "temp.vbs")
        
        ReDim Preserve asRepLine01(1)
        asRepLine01(0) = Chr(9) & "aaa"
        asRepLine01(1) = Chr(9) & "bbb"
        sKeyword = "'>>>インクルード<<<"
        bRet = ReplaceArray(sKeyword, asRepLine01, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "が見つかりませんでした"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine02(0)
        'asRepLine02(0) = ""
        sKeyword = "'>>>変数定義<<<"
        bRet = ReplaceArray(sKeyword, asRepLine02, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "が見つかりませんでした"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine03(3)
        asRepLine03(0) = Chr(9) & "ccc"
        asRepLine03(1) = Chr(9) & "dddddd"
        asRepLine03(2) = ""
        asRepLine03(3) = Chr(9) & "e"
        sKeyword = "'>>>本処理<<<"
        bRet = ReplaceArray(sKeyword, asRepLine03, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "が見つかりませんでした"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine04(0)
        asRepLine04(0) = "888888888888888"
        sKeyword = "'>>>関数定義<<<"
        bRet = ReplaceArray(sKeyword, asRepLine04, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "が見つかりませんでした"
            Debug.Assert 0
        End If
        
        'ReDim Preserve asRepLine05(0)
        'asRepLine05(0) = ""
        sKeyword = "'>>>定数定義<<<"
        bRet = ReplaceArray(sKeyword, asRepLine05, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "が見つかりませんでした"
            Debug.Assert 0
        End If
        
        Call OutputTxtFile(objWshShell.SpecialFolders("Desktop") & "\" & "temp2.vbs", asBaseFileLine)
    End Sub

' ==================================================================
' = 概要    セル範囲（Range型）を文字列配列（String配列型）に変換する。
' =         主にセル範囲をテキストファイルに出力する時に使用する。
' = 引数    rCellsRange             Range   [in]  対象のセル範囲
' = 引数    asLine()                String  [out] 文字列返還後のセル範囲
' = 引数    bIgnoreInvisibleCell    String  [in]  非表示セル無視実行可否
' = 引数    sDelimiter              String  [in]  区切り文字
' = 戻値    なし
' = 覚書    列が隣り合ったセル同士は指定された区切り文字で区切られる
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
Public Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIgnoreInvisibleCell As Boolean, _
    ByVal sDelimiter As String _
)
    Dim lLineIdx As Long
    lLineIdx = 0
    ReDim Preserve asLine(lLineIdx)
    
    Dim lRowIdx As Long
    For lRowIdx = 1 To rCellsRange.Rows.Count
        Dim lIgnoreCnt As Long
        lIgnoreCnt = 0
        Dim lClmIdx As Long
        For lClmIdx = 1 To rCellsRange.Columns.Count
            Dim sCurCellValue As String
            sCurCellValue = rCellsRange(lRowIdx, lClmIdx).Value
            '非表示セルは無視する
            Dim bIsIgnoreCurExec As Boolean
            If bIgnoreInvisibleCell = True Then
                If rCellsRange(lRowIdx, lClmIdx).EntireRow.Hidden = True Or _
                   rCellsRange(lRowIdx, lClmIdx).EntireColumn.Hidden = True Then
                    bIsIgnoreCurExec = True
                Else
                    bIsIgnoreCurExec = False
                End If
            Else
                bIsIgnoreCurExec = False
            End If
            
            If bIsIgnoreCurExec = True Then
                lIgnoreCnt = lIgnoreCnt + 1
            Else
                If lClmIdx = 1 Then
                    asLine(lLineIdx) = sCurCellValue
                Else
                    asLine(lLineIdx) = asLine(lLineIdx) & sDelimiter & sCurCellValue
                End If
            End If
        Next lClmIdx
        If lIgnoreCnt = rCellsRange.Columns.Count Then '非表示行は行加算しない
            'Do Nothing
        Else
            If lRowIdx = rCellsRange.Rows.Count Then '最終行は行加算しない
                'Do Nothing
            Else
                lLineIdx = lLineIdx + 1
                ReDim Preserve asLine(lLineIdx)
            End If
        End If
    Next lRowIdx
End Function
    Private Sub Test_ConvRange2Array()
        '★
    End Sub

