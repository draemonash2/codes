Attribute VB_Name = "Mng_ExcelOpe"
Option Explicit

' excel operation library v2.3

Sub CreateSheetList()
    Dim oSheet As Object
    Dim lRowIdx As Long
    Dim lColumnIdx As Long
 
    If MsgBox("アクティブセルから下にシート名一覧を作成してもいいですか？", vbYesNo + vbDefaultButton2) = vbNo Then
        'None
    Else
        lRowIdx = ActiveCell.Row
        lColumnIdx = ActiveCell.Column
 
        For Each oSheet In ActiveWorkbook.Sheets
            Cells(lRowIdx, lColumnIdx).Value = oSheet.Name
            lRowIdx = lRowIdx + 1
        Next oSheet
    End If
End Sub

' ============================================
' = 概要    アクティブシートの B2 セル以下に記述された関数名を基に
' =         関数ツリーを作成するためのオブジェクトを生成する。
' =         関数名が記述されたテキストボックスと左記に接続されたコネクタを、
' =         関数名の数だけ生成する。
' = 覚書    なし
' ============================================
Sub CreateFuncTree()
    ' 変数定義
    Dim intHeight As Integer                ' 追加するテキストボックスの位置(高さ)基準
    Dim intStartRow As Integer              ' スタートする行数
    Dim intConnecterBeginPointY As Integer  ' コネクタ始点の垂直位置
    Dim intConnecterBeginPointX As Integer  ' コネクタ始点の水平位置
    Dim shpObjectBox As Shape               ' 関数名ボックス定義
    Dim shpObjectLine As Shape              ' コネクタ定義
 
    ' 関数名ボックス生成位置定義
    intHeight = 25
    intStartRow = 2
 
    For intSearchRow = intStartRow To (Range("B2").End(xlDown).Row)
        ' === 関数名ボックス生成 ===
            ' オブジェクト生成
            Set shpObjectBox = ActiveSheet.Shapes.AddShape( _
                Type:=msoShapeFlowchartPredefinedProcess, _
                Left:=200, _
                Top:=intHeight * (intSearchRow - 1), _
                Width:=100, _
                Height:=100)
 
            ' 書式設定
            shpObjectBox.Fill.ForeColor.RGB = RGB(128, 0, 0)     ' 背景色
            shpObjectBox.Line.ForeColor.RGB = RGB(0, 0, 0)       ' 線の色
            shpObjectBox.Line.Weight = 2                         ' 線の太さ
            shpObjectBox.Select
            Selection.Characters.Text = Cells(intSearchRow, 2)   ' テキストに数式の内容を設定
            Selection.AutoSize = True                            ' 自動サイズ調整にする
 
        ' === コネクタ生成 ===
            ' オブジェクト生成
            intConnecterBeginPointY = (intHeight * (intSearchRow - 1)) - 10
            intConnecterBeginPointX = 200
            Set shpObjectLine = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, intConnecterBeginPointX, intConnecterBeginPointY, 0, 0)
            ' 書式設定
            shpObjectLine.Line.ForeColor.RGB = RGB(0, 0, 0)      ' 線の色
            shpObjectLine.Line.Weight = 2                        ' 線の太さ
            ' コネクタ接続
            shpObjectLine.Select
            Selection.ShapeRange.ConnectorFormat.EndConnect shpObjectBox, 1
 
   Next intSearchRow
End Sub

' ==================================================================
' = 概要    指定した範囲の文字列を結合する
' =         区切り文字を指定した場合、結合する間に文字を挿入する
' = 引数    rConcRange    Range   [in]  結合する範囲
' = 引数    sDlmtr        String  [in]  区切り文字
' = 戻値                  Variant       結合後の文字列
' = 覚書    なし
' ==================================================================
Public Function ConcStr( _
    ByRef rConcRange As Range, _
    Optional ByVal sDlmtr As String _
) As Variant
    Dim rConcRangeCnt As Range
    Dim sConcTxtBuf As String
    
    If rConcRange Is Nothing Then
        ConcStr = CVErr(xlErrRef)  'エラー値
    Else
        If rConcRange.Rows.Count = 1 Or _
           rConcRange.Columns.Count = 1 Then
            For Each rConcRangeCnt In rConcRange
                sConcTxtBuf = sConcTxtBuf & sDlmtr & rConcRangeCnt.Value
            Next rConcRangeCnt
            
            ' 区切り文字判定
            If sDlmtr <> "" Then
                ConcStr = Mid$(sConcTxtBuf, Len(sDlmtr) + 1)
            Else
                ConcStr = sConcTxtBuf
            End If
        Else
            ConcStr = CVErr(xlErrRef)  'エラー値
        End If
    End If
End Function
    Private Sub Test_ConcStr()
        Dim oTrgtRangePos01 As Range
        Dim oTrgtRangePos02 As Range
        Dim oTrgtRangePos03 As Range
        Dim oTrgtRangeNeg01 As Range
        Dim oTrgtRangeNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtRangePos01 = .Cells(1, 1)
            Set oTrgtRangePos02 = .Range(.Cells(1, 1), .Cells(1, 3))
            Set oTrgtRangePos03 = .Range(.Cells(1, 3), .Cells(1, 1))
            Set oTrgtRangeNeg01 = .Range(.Cells(1, 1), .Cells(3, 3))
            Set oTrgtRangeNeg02 = Nothing
        End With
        
        Dim lIdx
        Dim asStrBefore() As String
        For lIdx = 0 To oTrgtRangeNeg01.Count
            ReDim Preserve asStrBefore(lIdx)
            asStrBefore(lIdx) = oTrgtRangeNeg01.Item(lIdx + 1)
        Next lIdx
        
        Debug.Print "*** test start! ***"
        oTrgtRangePos01.Item(1) = "aaa"
        Debug.Print ConcStr(oTrgtRangePos01, "\") 'aaa
        Debug.Print ConcStr(oTrgtRangePos01, "")  'aaa
        oTrgtRangePos02.Item(1) = "bbb"
        oTrgtRangePos02.Item(2) = "ccc"
        oTrgtRangePos02.Item(3) = "ddd"
        Debug.Print ConcStr(oTrgtRangePos02, "\")  'bbb\ccc\ddd
        Debug.Print ConcStr(oTrgtRangePos02, "  ") 'bbb  ccc  ddd
        Debug.Print ConcStr(oTrgtRangePos02, "")   'bbbcccddd
        oTrgtRangePos03.Item(1) = "eee"
        oTrgtRangePos03.Item(2) = "fff"
        oTrgtRangePos03.Item(3) = "ggg"
        Debug.Print ConcStr(oTrgtRangePos03, "\")  'eee\fff\ggg
        Debug.Print ConcStr(oTrgtRangePos03, "  ") 'eee  fff  ggg
        Debug.Print ConcStr(oTrgtRangePos03, "")   'eeefffggg
        Debug.Print ConcStr(oTrgtRangeNeg01, "\")  'エラー 2023
        Debug.Print ConcStr(oTrgtRangeNeg02, "\")  'エラー 2023
        Debug.Print "*** test finished! ***"
        
        For lIdx = 0 To oTrgtRangeNeg01.Count
            oTrgtRangeNeg01.Item(lIdx + 1) = asStrBefore(lIdx)
        Next lIdx
    End Sub

' ==================================================================
' = 概要    文字列を分割し、指定した要素の文字列を返却する
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 引数    iExtIndex   String  [in]  抽出する要素 ( 0 origin )
' = 戻値                Variant       抽出文字列
' = 覚書    iExtIndex が要素を超える場合、空文字列を返却する
' ==================================================================
Public Function SplitStr( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As Variant
    If sDlmtr = "" Then
        SplitStr = sStr
    Else
        If sStr = "" Then
            SplitStr = ""
        Else
            Dim vSplitStr As Variant
            vSplitStr = Split(sStr, sDlmtr) ' 文字列分割
            If iExtIndex > UBound(vSplitStr) Or _
               iExtIndex < LBound(vSplitStr) Then
                SplitStr = ""
            Else
                SplitStr = vSplitStr(iExtIndex)
            End If
        End If
    End If
End Function
    Private Sub Test_SplitStr()
        Debug.Print "*** test start! ***"
        Debug.Print SplitStr("c:\test\a.txt", "\", 0)  'c:
        Debug.Print SplitStr("c:\test\a.txt", "\", 1)  'test
        Debug.Print SplitStr("c:\test\a.txt", "\", 2)  'a.txt
        Debug.Print SplitStr("c:\test\a.txt", "\", -1) '
        Debug.Print SplitStr("c:\test\a.txt", "\", 3)  '
        Debug.Print SplitStr("", "\", 1)               '
        Debug.Print SplitStr("c:\a.txt", "", 1)        'c:\a.txt
        Debug.Print SplitStr("", "", 1)                '
        Debug.Print SplitStr("", "", 0)                '
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    取り消し線の有無を判定する (TRUE:有、FALSE:無)
' = 引数    rRange   Range     [in]  セル
' = 戻値             Variant         取り消し線有無
' = 覚書    なし
' ==================================================================
Public Function GetStrikeExist( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetStrikeExist = rRange.Font.Strikethrough
    Else
        GetStrikeExist = CVErr(xlErrRef)  'エラー値
    End If
End Function
    Private Sub Test_GetStrikeExist()
        Dim oTrgtRangePos01 As Range
        Dim oTrgtRangeNeg01 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtRangePos01 = .Cells(1, 1)
            Set oTrgtRangeNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
        End With
        
        Dim lStrkThrghBefore As Long
        lStrkThrghBefore = oTrgtRangePos01.Font.Strikethrough
        
        Debug.Print "*** test start! ***"
        oTrgtRangePos01.Font.Strikethrough = True
        Debug.Print GetStrikeExist(oTrgtRangePos01) 'True
        oTrgtRangePos01.Font.Strikethrough = False
        Debug.Print GetStrikeExist(oTrgtRangePos01) 'False
        Debug.Print GetStrikeExist(oTrgtRangeNeg01) 'エラー 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtRangePos01.Font.Strikethrough = lStrkThrghBefore
    End Sub

' ==================================================================
' = 概要    フォントカラーを返却する
' = 引数    rTrgtRange  Range     [in]  セル
' = 引数    sColorType  String    [in]  色種別（R or G or B）
' = 引数    bIsHex      Boolean   [in]  基数（0:Decimal、1:Hex）
' = 戻値                Variant         フォント色
' = 覚書    なし
' ==================================================================
Public Function GetFontColor( _
    ByRef rTrgtCell As Range, _
    ByVal sColorType As String, _
    ByVal bIsHex As Boolean _
) As Variant
    Dim lColorRGB As Long
    Dim lColorX As Long
    
    If rTrgtCell.Count > 1 Then
        GetFontColor = CVErr(xlErrRef)
    Else
        lColorRGB = rTrgtCell.Font.Color
        lColorX = ConvRgb2X(lColorRGB, sColorType)
        If lColorX > 255 Then
            GetFontColor = CVErr(xlErrValue)
        Else
            If bIsHex = True Then
                GetFontColor = UCase(String(2 - Len(Hex(lColorX)), "0") & Hex(lColorX))
            Else
                GetFontColor = lColorX
            End If
        End If
    End If
End Function
    Private Sub Test_GetFontColor()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Font.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Font.Color = RGB(0, 0, 0)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '0
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '0
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '0
        oTrgtCellsPos01.Font.Color = RGB(100, 100, 100)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '100
        oTrgtCellsPos01.Font.Color = RGB(255, 255, 255)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '255
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '255
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '255
        oTrgtCellsPos01.Font.Color = RGB(16, 100, 152)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '16
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '152
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", True)   '10
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", True)   '64
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", True)   '98
        Debug.Print GetFontColor(oTrgtCellsPos01, "", False)   'エラー 2015
        Debug.Print GetFontColor(oTrgtCellsPos01, "aa", False) 'エラー 2015
        Debug.Print GetFontColor(oTrgtCellsNeg01, "r", False)  'エラー 2023
        Debug.Print GetFontColor(oTrgtCellsNeg02, "r", False)  'エラー 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Font.Color = lColorBefore
    End Sub

' ==================================================================
' = 概要    背景色を返却する
' = 引数    rTrgtRange  Range     [in]  セル
' = 引数    sColorType  String    [in]  色種別（R or G or B）
' = 引数    bIsHex      Boolean   [in]  基数（0:Decimal、1:Hex）
' = 戻値                Variant         背景色
' = 覚書    なし
' ==================================================================
Public Function GetInteriorColor( _
    ByRef rTrgtCell As Range, _
    ByVal sColorType As String, _
    ByVal bIsHex As Boolean _
) As Variant
    Dim lColorRGB As Long
    Dim lColorX As Long
    
    If rTrgtCell.Count > 1 Then
        GetInteriorColor = CVErr(xlErrRef)
    Else
        lColorRGB = rTrgtCell.Interior.Color
        lColorX = ConvRgb2X(lColorRGB, sColorType)
        If lColorX > 255 Then
            GetInteriorColor = CVErr(xlErrValue)
        Else
            If bIsHex = True Then
                GetInteriorColor = UCase(String(2 - Len(Hex(lColorX)), "0") & Hex(lColorX))
            Else
                GetInteriorColor = lColorX
            End If
        End If
    End If
End Function
    Private Sub Test_GetInteriorColor()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Interior.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Interior.Color = RGB(0, 0, 0)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '0
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '0
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '0
        oTrgtCellsPos01.Interior.Color = RGB(100, 100, 100)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '100
        oTrgtCellsPos01.Interior.Color = RGB(255, 255, 255)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '255
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '255
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '255
        oTrgtCellsPos01.Interior.Color = RGB(16, 100, 152)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '16
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '152
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", True)   '10
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", True)   '64
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", True)   '98
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "", False)   'エラー 2015
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "aa", False) 'エラー 2015
        Debug.Print GetInteriorColor(oTrgtCellsNeg01, "r", False)  'エラー 2023
        Debug.Print GetInteriorColor(oTrgtCellsNeg02, "r", False)  'エラー 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Interior.Color = lColorBefore
    End Sub

' ==================================================================
' = 概要    ビットＡＮＤ演算を行う。（数値）
' = 引数    cInVal1   Currency   [in]  入力値 左項（10進数数値）
' = 引数    cInVal2   Currency   [in]  入力値 右項（10進数数値）
' = 戻値              Variant          演算結果（10進数数値）
' = 覚書    なし
' ==================================================================
Public Function BitAndVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    If cInVal1 > 2147483647# Or cInVal1 < -2147483647# Or _
       cInVal2 > 2147483647# Or cInVal2 < -2147483647# Then
        BitAndVal = CVErr(xlErrNum)  'エラー値
    Else
        BitAndVal = cInVal1 And cInVal2
    End If
End Function
    Private Sub Test_BitAndVal()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndVal(&HFFFF&, &HFF00&)         '65280 (0xFF00)
        Debug.Print BitAndVal(&HFFFF&, &HFF&)           '255 (0xFF)
        Debug.Print BitAndVal(&HFFFF&, &HA5A5&)         '42405 (0xA5A5)
        Debug.Print BitAndVal(&HA5&, &HA500&)           '0
        Debug.Print BitAndVal(&H1&, &H8&)               '0
        Debug.Print BitAndVal(&H1&, &HA&)               '0
        Debug.Print BitAndVal(&H5&, &HA&)               '0
        Debug.Print BitAndVal(&H7FFFFFFF, &HFF&)        '255 (0xFF)
        Debug.Print BitAndVal(&H80000000, &HFF&)        'エラー 2036
        Debug.Print BitAndVal(2147483648#, &HFF&)       'エラー 2036
        Debug.Print BitAndVal(2147483647#, &HFF&)       '255 (0xFF)
        Debug.Print BitAndVal(-2147483647#, &HFF&)      '1
        Debug.Print BitAndVal(-2147483648#, &HFF&)      'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＡＮＤ演算を行う。（文字列１６進数）
' = 引数    sInHexVal1  String     [in]  入力値 左項（文字列）
' = 引数    sInHexVal2  String     [in]  入力値 右項（文字列）
' = 引数    lInDigitNum Long       [in]  出力桁数
' = 戻値                Variant          演算結果（文字列）
' = 覚書    なし
' ==================================================================
Public Function BitAndStrHex( _
    ByVal sInHexVal1 As String, _
    ByVal sInHexVal2 As String, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal1 = "" Or sInHexVal2 = "" Then
        BitAndStrHex = CVErr(xlErrNull) 'エラー値
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitAndStrHex = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    '入力値１のＢＩＮ変換
    Dim sInBinVal1 As String
    sInBinVal1 = Hex2Bin(sInHexVal1)
    If sInBinVal1 = "error" Then
        BitAndStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal1 : " & sInBinVal1
    
    '入力値２のＢＩＮ変換
    Dim sInBinVal2 As String
    sInBinVal2 = Hex2Bin(sInHexVal2)
    If sInBinVal2 = "error" Then
        BitAndStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal2 : " & sInBinVal2
    
    'ＢＩＮ ＡＮＤ演算
    Dim sOutBinVal As String
    sOutBinVal = BitAndStrBin(sInBinVal1, sInBinVal2, lInDigitNum * 4)
    
    'ＢＩＮ⇒ＨＥＸ変換
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitAndStrHex = sOutHexVal
    
End Function
    Private Sub Test_BitAndStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndStrHex("FF", "FF00")                  '0000
        Debug.Print BitAndStrHex("A5A5", "5A5A")                '0000
        Debug.Print BitAndStrHex("A5A5", "00FF")                '00A5
        Debug.Print BitAndStrHex("A5", "00FF")                  '00A5
        Debug.Print BitAndStrHex("FFFF0B00", "01010300")        '01010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 10)    '0001010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 8)     '01010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 7)     '1010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 6)     '010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 5)     '10300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 4)     '0300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 2)     '00
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 1)     '0
        Debug.Print BitAndStrHex("ab", "00FF")                  '00AB
        Debug.Print BitAndStrHex("cd", "00FF")                  '00CD
        Debug.Print BitAndStrHex("ef", "00FF")                  '00EF
        Debug.Print BitAndStrHex(" 0B00", "0300")               'エラー 2015
        Debug.Print BitAndStrHex("", "0300")                    'エラー 2000
        Debug.Print BitAndStrHex("0B00", "")                    'エラー 2000
        Debug.Print BitAndStrHex("0B00", "0300", -1)            'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＡＮＤ演算を行う。（文字列２進数）
' = 引数    sInBinVal1  String     [in]  入力値 左項（文字列）
' = 引数    sInBinVal2  String     [in]  入力値 右項（文字列）
' = 引数    lInBitLen   Long       [in]  出力ビット数
' = 戻値                Variant          演算結果（文字列）
' = 覚書    なし
' ==================================================================
Public Function BitAndStrBin( _
    ByVal sInBinVal1 As String, _
    ByVal sInBinVal2 As String, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal1 = "" Or sInBinVal2 = "" Then
        BitAndStrBin = CVErr(xlErrNum) 'エラー値
    Else
        '文字数を合わせる
        If Len(sInBinVal1) > Len(sInBinVal2) Then
            sInBinVal2 = String(Len(sInBinVal1) - Len(sInBinVal2), "0") & sInBinVal2
        Else
            sInBinVal1 = String(Len(sInBinVal2) - Len(sInBinVal1), "0") & sInBinVal1
        End If
        Debug.Assert Len(sInBinVal1) = Len(sInBinVal2)
        
        'OR演算
        Dim lValIdx As Long
        Dim sOutBin As String
        Dim bIsError As Boolean
        lValIdx = Len(sInBinVal1)
        sOutBin = ""
        bIsError = False
        Do
            Select Case Mid$(sInBinVal1, lValIdx, 1) & Mid$(sInBinVal2, lValIdx, 1)
                Case "00": sOutBin = "0" & sOutBin
                Case "10": sOutBin = "0" & sOutBin
                Case "01": sOutBin = "0" & sOutBin
                Case "11": sOutBin = "1" & sOutBin
                Case Else: bIsError = True
            End Select
            lValIdx = lValIdx - 1
        Loop While lValIdx > 0 And bIsError = False
        
        If bIsError = True Then
            BitAndStrBin = CVErr(xlErrNum) 'エラー値
        Else
            If lInBitLen = 0 Then
                BitAndStrBin = sOutBin
            Else
                If lInBitLen <= Len(sOutBin) Then
                    BitAndStrBin = Right$(sOutBin, lInBitLen)
                Else
                    BitAndStrBin = String(lInBitLen - Len(sOutBin), "0") & sOutBin
                End If
            End If
        End If
    End If
End Function
    Private Sub Test_BitAndStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndStrBin("11", "11110")          '00010
        Debug.Print BitAndStrBin("11110", "11")          '00010
        Debug.Print BitAndStrBin("1", "1")               '1
        Debug.Print BitAndStrBin("1", "0")               '0
        Debug.Print BitAndStrBin("0", "1")               '0
        Debug.Print BitAndStrBin("0", "0")               '0
        Debug.Print BitAndStrBin("00000011", "11000000") '00000000
        Debug.Print BitAndStrBin("0111", "0010", 10)     '0000000010
        Debug.Print BitAndStrBin("0111", "0010", 0)      '0010
        Debug.Print BitAndStrBin("0111", "0010", 2)      '10
        Debug.Print BitAndStrBin("0111", "0010", 1)      '0
        Debug.Print BitAndStrBin("0101", "001F")         'エラー 2036
        Debug.Print BitAndStrBin(" 101", "0010")         'エラー 2036
        Debug.Print BitAndStrBin("", "0010")             'エラー 2036
        Debug.Print BitAndStrBin("0101", "")             'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＯＲ演算を行う。（数値）
' = 引数    cInVal1   Currency   [in]  入力値 左項（10進数数値）
' = 引数    cInVal2   Currency   [in]  入力値 右項（10進数数値）
' = 戻値              Variant          演算結果（10進数数値）
' = 覚書    なし
' ==================================================================
Public Function BitOrVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    Dim sHexVal As String
    If cInVal1 > 2147483647# Or cInVal1 < -2147483647# Or _
       cInVal2 > 2147483647# Or cInVal2 < -2147483647# Then
        BitOrVal = CVErr(xlErrNum)  'エラー値
    Else
        BitOrVal = cInVal1 Or cInVal2
    End If
End Function
    Private Sub Test_BitOrVal()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrVal(&HFFFF&, &HFF00&)      '65535 (0xFFFF)
        Debug.Print BitOrVal(&HFFFF&, &HFF&)        '65535 (0xFFFF)
        Debug.Print BitOrVal(&HFFFF&, &HA5A5&)      '65535 (0xFFFF)
        Debug.Print BitOrVal(&HA5&, &HA500&)        '42405 (0xA5A5)
        Debug.Print BitOrVal(&H1&, &H8&)            '9
        Debug.Print BitOrVal(&H1&, &HA&)            '11 (0xB)
        Debug.Print BitOrVal(&H5&, &HA&)            '15 (0xF)
        Debug.Print BitOrVal(&H7FFFFFFF, &HFF&)     '2147483647 (0x7FFFFFFF)
        Debug.Print BitOrVal(&H80000000, &HFF&)     'エラー 2036
        Debug.Print BitOrVal(2147483648#, &HFF&)    'エラー 2036
        Debug.Print BitOrVal(2147483647#, &HFF&)    '2147483647 (0x7FFFFFFF)
        Debug.Print BitOrVal(-2147483647#, &HFF&)   '-2147483393 (0x800000FF)
        Debug.Print BitOrVal(-2147483648#, &HFF&)   'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＯＲ演算を行う。（文字列１６進数）
' = 引数    sInHexVal1  String     [in]  入力値 左項（文字列）
' = 引数    sInHexVal2  String     [in]  入力値 右項（文字列）
' = 引数    lInDigitNum Long       [in]  出力桁数
' = 戻値                Variant          演算結果（文字列）
' = 覚書    なし
' ==================================================================
Public Function BitOrStrHex( _
    ByVal sInHexVal1 As String, _
    ByVal sInHexVal2 As String, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal1 = "" Or sInHexVal2 = "" Then
        BitOrStrHex = CVErr(xlErrNull) 'エラー値
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitOrStrHex = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    '入力値１のＢＩＮ変換
    Dim sInBinVal1 As String
    sInBinVal1 = Hex2Bin(sInHexVal1)
    If sInBinVal1 = "error" Then
        BitOrStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal1 : " & sInBinVal1
    
    '入力値２のＢＩＮ変換
    Dim sInBinVal2 As String
    sInBinVal2 = Hex2Bin(sInHexVal2)
    If sInBinVal2 = "error" Then
        BitOrStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal2 : " & sInBinVal2
    
    'ＢＩＮ ＯＲ演算
    Dim sOutBinVal As String
    sOutBinVal = BitOrStrBin(sInBinVal1, sInBinVal2, lInDigitNum * 4)
    
    'ＢＩＮ⇒ＨＥＸ変換
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitOrStrHex = sOutHexVal
End Function
    Private Sub Test_BitOrStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrStrHex("FF", "FF00")                  'FFFF
        Debug.Print BitOrStrHex("A5A5", "5A5A")                'FFFF
        Debug.Print BitOrStrHex("A5A5", "00FF")                'A5FF
        Debug.Print BitOrStrHex("A5", "00FF")                  '00FF
        Debug.Print BitOrStrHex("FFFF0800", "01010300")        'FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 10)    '00FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 8)     'FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 7)     'FFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 6)     'FF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 5)     'F0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 4)     '0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 2)     '00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 1)     '0
        Debug.Print BitOrStrHex("ab", "0000")                  '00AB
        Debug.Print BitOrStrHex("cd", "0000")                  '00CD
        Debug.Print BitOrStrHex("ef", "0000")                  '00EF
        Debug.Print BitOrStrHex(" 0800", "0300")               'エラー 2015
        Debug.Print BitOrStrHex("", "0300")                    'エラー 2000
        Debug.Print BitOrStrHex("0800", "")                    'エラー 2000
        Debug.Print BitOrStrHex("0800", "0300", -1)            'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＯＲ演算を行う。（文字列２進数）
' = 引数    sInBinVal1  String     [in]  入力値 左項（文字列）
' = 引数    sInBinVal2  String     [in]  入力値 右項（文字列）
' = 引数    lInBitLen   Long       [in]  出力ビット数
' = 戻値                Variant          演算結果（文字列）
' = 覚書    なし
' ==================================================================
Public Function BitOrStrBin( _
    ByVal sInBinVal1 As String, _
    ByVal sInBinVal2 As String, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal1 = "" Or sInBinVal2 = "" Then
        BitOrStrBin = CVErr(xlErrNum) 'エラー値
    Else
        '文字数を合わせる
        If Len(sInBinVal1) > Len(sInBinVal2) Then
            sInBinVal2 = String(Len(sInBinVal1) - Len(sInBinVal2), "0") & sInBinVal2
        Else
            sInBinVal1 = String(Len(sInBinVal2) - Len(sInBinVal1), "0") & sInBinVal1
        End If
        Debug.Assert Len(sInBinVal1) = Len(sInBinVal2)
        
        'OR演算
        Dim lValIdx As Long
        Dim sOutBin As String
        Dim bIsError As Boolean
        lValIdx = Len(sInBinVal1)
        sOutBin = ""
        bIsError = False
        Do
            Select Case Mid$(sInBinVal1, lValIdx, 1) & Mid$(sInBinVal2, lValIdx, 1)
                Case "00": sOutBin = "0" & sOutBin
                Case "10": sOutBin = "1" & sOutBin
                Case "01": sOutBin = "1" & sOutBin
                Case "11": sOutBin = "1" & sOutBin
                Case Else: bIsError = True
            End Select
            lValIdx = lValIdx - 1
        Loop While lValIdx > 0 And bIsError = False
        
        If bIsError = True Then
            BitOrStrBin = CVErr(xlErrNum) 'エラー値
        Else
            If lInBitLen = 0 Then
                BitOrStrBin = sOutBin
            Else
                If lInBitLen <= Len(sOutBin) Then
                    BitOrStrBin = Right$(sOutBin, lInBitLen)
                Else
                    BitOrStrBin = String(lInBitLen - Len(sOutBin), "0") & sOutBin
                End If
            End If
        End If
    End If
End Function
    Private Sub Test_BitOrStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrStrBin("11", "11110")                  '11111
        Debug.Print BitOrStrBin("11110", "11")                  '11111
        Debug.Print BitOrStrBin("1", "1")                       '1
        Debug.Print BitOrStrBin("1", "0")                       '1
        Debug.Print BitOrStrBin("0", "1")                       '1
        Debug.Print BitOrStrBin("0", "0")                       '0
        Debug.Print BitOrStrBin("00000011", "11000000")         '11000011
        Debug.Print BitOrStrBin("01010101", "00010010", 0)      '01010111
        Debug.Print BitOrStrBin("01010101", "00010010", 10)     '0001010111
        Debug.Print BitOrStrBin("01010101", "00010010", 7)      '1010111
        Debug.Print BitOrStrBin("01010101", "00010010", 6)      '010111
        Debug.Print BitOrStrBin("01010101", "00010010", 4)      '0111
        Debug.Print BitOrStrBin("01010101", "00010010", 3)      '111
        Debug.Print BitOrStrBin("01010101", "00010010", 2)      '11
        Debug.Print BitOrStrBin("01010101", "00010010", 1)      '1
        Debug.Print BitOrStrBin("0101", "001F")                 'エラー 2036
        Debug.Print BitOrStrBin(" 101", "0010")                 'エラー 2036
        Debug.Print BitOrStrBin("K01", "0010")                  'エラー 2036
        Debug.Print BitOrStrBin("", "0010")                     'エラー 2036
        Debug.Print BitOrStrBin("0101", "")                     'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＳＨＩＦＴ演算を行う。（数値）
' = 引数    cInDecVal       Currency  [in]  入力値（10進数数値）
' = 引数    lInShiftNum     Long      [in]  シフトビット数
' = 引数    eInDirection    Enum      [in]  シフト方向（0:左 1:右）
' = 引数    eInShiftType    Enum      [in]  シフト種別（0:論理 1:算術）
' = 戻値                    Variant         シフト結果（10進数数値）
' = 覚書    32ビットのみ対応する。そのため、左シフトの結果が32ビットを
' =         超える場合、下位32ビットのシフト結果を返却する。
' ==================================================================
Public Function BitShiftVal( _
    ByVal cInDecVal As Currency, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT _
) As Variant
    If cInDecVal < -2147483648# Or cInDecVal > 4294967295# Then
        BitShiftVal = CVErr(xlErrNum)  'エラー値
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftVal = CVErr(xlErrNum)  'エラー値
        Exit Function
    End If
    If eInDirection <> RIGHT_SHIFT And eInDirection <> LEFT_SHIFT Then
        BitShiftVal = CVErr(xlErrValue)  'エラー値
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftVal = CVErr(xlErrValue)  'エラー値
        Exit Function
    End If
    
    'Dec⇒Hex
    Dim sPreHexVal As String
    sPreHexVal = Dec2Hex(cInDecVal)
    If sPreHexVal = "error" Then
        BitShiftVal = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    'Hex⇒Bin
    Dim sPreBinVal As String
    sPreBinVal = Hex2Bin(sPreHexVal)
    Debug.Assert sPreBinVal <> "error"
    
    'Shift
    Dim sPostBinVal As String
    sPostBinVal = BitShiftStrBin(sPreBinVal, lInShiftNum, eInDirection, eInShiftType, 32)
    Debug.Assert sPostBinVal <> "error"
    
    'Bin⇒Hex
    Dim sPostHexVal As String
    sPostHexVal = Bin2Hex(sPostBinVal, True)
    Debug.Assert sPostHexVal <> "error"
    
    'Hex⇒Dec
    Dim vOutDecVal As Variant
    vOutDecVal = Hex2Dec(sPostHexVal, False)
    If vOutDecVal = "error" Then
        BitShiftVal = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    BitShiftVal = vOutDecVal
End Function
    Private Sub Test_BitShiftVal()
        Debug.Print "*** test start! ***"
        Debug.Print Hex(BitShiftVal(&H10&, 0, RIGHT_SHIFT, LOGICAL_SHIFT))          '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, RIGHT_SHIFT, LOGICAL_SHIFT))          '8
        Debug.Print Hex(BitShiftVal(&H10&, 2, RIGHT_SHIFT, LOGICAL_SHIFT))          '4
        Debug.Print Hex(BitShiftVal(&H10&, 3, RIGHT_SHIFT, LOGICAL_SHIFT))          '2
        Debug.Print Hex(BitShiftVal(&H10&, 4, RIGHT_SHIFT, LOGICAL_SHIFT))          '1
        Debug.Print Hex(BitShiftVal(&H10&, 5, RIGHT_SHIFT, LOGICAL_SHIFT))          '0
        Debug.Print Hex(BitShiftVal(&H10&, 8, RIGHT_SHIFT, LOGICAL_SHIFT))          '0
        Debug.Print Hex(BitShiftVal(&H10&, 0, LEFT_SHIFT, LOGICAL_SHIFT))           '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, LEFT_SHIFT, LOGICAL_SHIFT))           '20
        Debug.Print Hex(BitShiftVal(&H10&, 2, LEFT_SHIFT, LOGICAL_SHIFT))           '40
        Debug.Print Hex(BitShiftVal(&H10&, 3, LEFT_SHIFT, LOGICAL_SHIFT))           '80
        Debug.Print Hex(BitShiftVal(&H10&, 8, LEFT_SHIFT, LOGICAL_SHIFT))           '1000
        Debug.Print Hex(BitShiftVal(&H10&, 12, LEFT_SHIFT, LOGICAL_SHIFT))          '10000
        Debug.Print Hex(BitShiftVal(&H10&, 16, LEFT_SHIFT, LOGICAL_SHIFT))          '100000
        Debug.Print Hex(BitShiftVal(&H10&, 20, LEFT_SHIFT, LOGICAL_SHIFT))          '1000000
        Debug.Print Hex(BitShiftVal(&H10&, 24, LEFT_SHIFT, LOGICAL_SHIFT))          '10000000
        Debug.Print Hex(BitShiftVal(&H10&, 25, LEFT_SHIFT, LOGICAL_SHIFT))          '20000000
        Debug.Print Hex(BitShiftVal(&H10&, 26, LEFT_SHIFT, LOGICAL_SHIFT))          '40000000
       'Debug.Print Hex(BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT))          'エラー（Hex()にてオーバーフロー）
        Debug.Print BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT)               '2147483648
        Debug.Print BitShiftVal(&H10&, 28, LEFT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print BitShiftVal(&H10&, 29, LEFT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print Hex(BitShiftVal(&H7FFFFFFF, 0, LEFT_SHIFT, LOGICAL_SHIFT))      '7FFFFFFF
        Debug.Print BitShiftVal(&H80000000, 0, LEFT_SHIFT, LOGICAL_SHIFT)           '2147483648 (0x80000000)
        Debug.Print BitShiftVal(4294967294#, 0, LEFT_SHIFT, LOGICAL_SHIFT)          '4294967294 (0xFFFFFFFE)
        Debug.Print BitShiftVal(&HFFFFFFFE, 0, LEFT_SHIFT, LOGICAL_SHIFT)           '4294967294 (0xFFFFFFFE)
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE) '32
        Debug.Print BitShiftVal(&H10&, -1, LEFT_SHIFT, LOGICAL_SHIFT)               'エラー 2036
        Debug.Print BitShiftVal(&H10&, 1, 3, LOGICAL_SHIFT)                         'エラー 2015
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, 3)                            'エラー 2015
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＳＨＩＦＴ演算を行う。（文字列１６進数）
' = 引数    sInHexVal       String  [in]    入力値（文字列）
' = 引数    lInShiftNum     Long    [in]    シフトビット数
' = 引数    eInDirection    Enum    [in]    シフト方向（0:左 1:右）
' = 引数    eInShiftType    Enum    [in]    シフト種別
' =                                           0:論理
' =                                           1:算術（符号ビット保持）(※1)
' =                                           2:算術（符号ビット切捨）(※1)
' = 引数    lInDigitNum     Long    [in]    出力桁数
' = 戻値                    Variant         シフト結果（文字列）
' = 覚書    (※1) 出力桁数が入力値（文字列）の長さよりも小さい場合に、
' =               符号ビットを保持するか、無視して切り捨てるかを選択する。
' =           ex1) 10101011 を出力桁数4として右1算術(符号ビット保持)シフト ⇒ 1101
' =           ex2) 10101011 を出力桁数4として右1算術(符号ビット切捨)シフト ⇒ 0101
' ==================================================================
Public Function BitShiftStrHex( _
    ByVal sInHexVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal = "" Then
        BitShiftStrHex = CVErr(xlErrNull) 'エラー値
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftStrHex = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    End If
    If eInDirection <> LEFT_SHIFT And eInDirection <> RIGHT_SHIFT Then
        BitShiftStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitShiftStrHex = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    '入力値のＨＥＸ⇒ＢＩＮ変換
    Dim sInBinVal As String
    sInBinVal = Hex2Bin(sInHexVal)
    If sInBinVal = "error" Then
        BitShiftStrHex = CVErr(xlErrValue) 'エラー値
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal : " & sInBinVal
    
    'ＢＩＮシフト
    Dim sOutBinVal As String
    Dim sTmpBinVal As String
    sTmpBinVal = BitShiftStrBin(sInBinVal, lInShiftNum, eInDirection, eInShiftType, lInDigitNum * 4)
    Dim lModNum As Long
    lModNum = Len(sTmpBinVal) Mod 4
    If lModNum = 0 Then
        sOutBinVal = sTmpBinVal
    Else
        sOutBinVal = String(4 - lModNum, "0") & sTmpBinVal
    End If
    
    'ＢＩＮ⇒ＨＥＸ変換
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitShiftStrHex = sOutHexVal
End Function
    Private Sub Test_BitShiftStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftStrHex("B0", 0, LEFT_SHIFT, LOGICAL_SHIFT)                  'B0
        Debug.Print BitShiftStrHex("B0", 1, LEFT_SHIFT, LOGICAL_SHIFT)                  '160
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT)                  '2C0
        Debug.Print BitShiftStrHex("B0", 3, LEFT_SHIFT, LOGICAL_SHIFT)                  '580
        Debug.Print BitShiftStrHex("B0", 4, LEFT_SHIFT, LOGICAL_SHIFT)                  'B00
        Debug.Print BitShiftStrHex("B0", 120, LEFT_SHIFT, LOGICAL_SHIFT)                'B0 + 0×30個
        Debug.Print BitShiftStrHex("B0", 0, RIGHT_SHIFT, LOGICAL_SHIFT)                 'B0
        Debug.Print BitShiftStrHex("B0", 1, RIGHT_SHIFT, LOGICAL_SHIFT)                 '58
        Debug.Print BitShiftStrHex("B0", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                 '2C
        Debug.Print BitShiftStrHex("B0", 3, RIGHT_SHIFT, LOGICAL_SHIFT)                 '16
        Debug.Print BitShiftStrHex("B0", 4, RIGHT_SHIFT, LOGICAL_SHIFT)                 'B
        Debug.Print BitShiftStrHex("B0", 7, RIGHT_SHIFT, LOGICAL_SHIFT)                 '1
        Debug.Print BitShiftStrHex("B0", 8, RIGHT_SHIFT, LOGICAL_SHIFT)                 '0
        Debug.Print BitShiftStrHex("B0", 9, RIGHT_SHIFT, LOGICAL_SHIFT)                 '0
        Debug.Print BitShiftStrHex("B0", 120, RIGHT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 0)               '2C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 1)               '0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 2)               'C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 3)               '2C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 4)               '02C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)              '00000002C0
        Debug.Print BitShiftStrHex("B0", 9, RIGHT_SHIFT, LOGICAL_SHIFT, 8)              '00000000
        Debug.Print BitShiftStrHex("B0", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE)   '160
        Debug.Print BitShiftStrHex("ab", 8, LEFT_SHIFT)                                 'AB00
        Debug.Print BitShiftStrHex("cd", 8, LEFT_SHIFT)                                 'CD00
        Debug.Print BitShiftStrHex("ef", 8, LEFT_SHIFT)                                 'EF00
        Debug.Print BitShiftStrHex("", 2, LEFT_SHIFT, LOGICAL_SHIFT)                    'エラー 2000
        Debug.Print BitShiftStrHex(" B", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                 'エラー 2015
        Debug.Print BitShiftStrHex("K", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                  'エラー 2015
        Debug.Print BitShiftStrHex("B", -1, LEFT_SHIFT, LOGICAL_SHIFT)                  'エラー 2036
        Debug.Print BitShiftStrHex("B", 1, 3, LOGICAL_SHIFT)                            'エラー 2015
        Debug.Print BitShiftStrHex("B", 1, LEFT_SHIFT, 3)                               'エラー 2015
        Debug.Print BitShiftStrHex("B", 1, LEFT_SHIFT, , -1)                            'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＳＨＩＦＴ演算を行う。（文字列２進数）
' = 引数    sInBinVal       String  [in]    入力値（文字列）
' = 引数    lInShiftNum     Long    [in]    シフトビット数
' = 引数    eInDirection    Enum    [in]    シフト方向（0:左 1:右）
' = 引数    eInShiftType    Enum    [in]    シフト種別
' =                                           0:論理
' =                                           1:算術（符号ビット保持）(※1)
' =                                           2:算術（符号ビット切捨）(※1)
' = 引数    lInBitLen       Long    [in]    出力ビット数
' = 戻値                    Variant         シフト結果（文字列）
' = 覚書    (※1) 出力桁数が入力値（文字列）の長さよりも小さい場合に、
' =               符号ビットを保持するか、無視して切り捨てるかを選択する。
' =           ex1) "AB"(0b10101011) を出力桁数1として右1算術(符号ビット保持)シフト
' =              ⇒ "D"(0b1101)
' =           ex2) "AB"(0b10101011) を出力桁数1として右1算術(符号ビット切捨)シフト
' =              ⇒ "5"(0b0101)
' ==================================================================
Public Function BitShiftStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal = "" Then
        BitShiftStrBin = CVErr(xlErrNull) 'エラー値
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftStrBin = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    If Replace(Replace(sInBinVal, "1", ""), "0", "") <> "" Then
        BitShiftStrBin = CVErr(xlErrValue) 'エラー値
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftStrBin = CVErr(xlErrValue) 'エラー値
        Exit Function
    End If
    If eInDirection <> LEFT_SHIFT And eInDirection <> RIGHT_SHIFT Then
        BitShiftStrBin = CVErr(xlErrValue) 'エラー値
        Exit Function
    End If
    If lInBitLen < 0 Then
        BitShiftStrBin = CVErr(xlErrNum) 'エラー値
        Exit Function
    End If
    
    Select Case eInShiftType
        Case LOGICAL_SHIFT:
            BitShiftStrBin = BitShiftLogStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen)
        Case ARITHMETIC_SHIFT_SIGNBITSAVE:
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen, True, False)
        Case ARITHMETIC_SHIFT_SIGNBITTRUNC:
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen, False, False)
        Case Else
            BitShiftStrBin = CVErr(xlErrValue) 'エラー値
    End Select
End Function
    Private Sub Test_BitShiftStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftStrBin("1011", 0, LEFT_SHIFT, LOGICAL_SHIFT)                        '1011
        Debug.Print BitShiftStrBin("1011", 1, LEFT_SHIFT, LOGICAL_SHIFT)                        '10110
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT)                        '101100
        Debug.Print BitShiftStrBin("1011", 4, LEFT_SHIFT, LOGICAL_SHIFT)                        '10110000
        Debug.Print BitShiftStrBin("1011", 0, RIGHT_SHIFT, LOGICAL_SHIFT)                       '1011
        Debug.Print BitShiftStrBin("1011", 1, RIGHT_SHIFT, LOGICAL_SHIFT)                       '101
        Debug.Print BitShiftStrBin("1011", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                       '10
        Debug.Print BitShiftStrBin("1011", 3, RIGHT_SHIFT, LOGICAL_SHIFT)                       '1
        Debug.Print BitShiftStrBin("1011", 4, RIGHT_SHIFT, LOGICAL_SHIFT)                       '0
        Debug.Print BitShiftStrBin("1011", 5, RIGHT_SHIFT, LOGICAL_SHIFT)                       '0
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 0)                     '101100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 2)                     '00
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 3)                     '100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 4)                     '1100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)                    '0000101100
        Debug.Print BitShiftStrBin("1011", 120, LEFT_SHIFT, LOGICAL_SHIFT)                      '1011 + 0×120個
        Debug.Print BitShiftStrBin("1011", 5, RIGHT_SHIFT, LOGICAL_SHIFT, 8)                    '00000000
        Debug.Print BitShiftStrBin("10001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16) '1111111000101100
        Debug.Print BitShiftStrBin("", 2, LEFT_SHIFT, LOGICAL_SHIFT)                            'エラー 2000
        Debug.Print BitShiftStrBin(":1011", 2, LEFT_SHIFT, LOGICAL_SHIFT)                       'エラー 2015
        Debug.Print BitShiftStrBin("1021", 2, LEFT_SHIFT, LOGICAL_SHIFT)                        'エラー 2015
        Debug.Print BitShiftStrBin("1011", -1, LEFT_SHIFT, LOGICAL_SHIFT)                       'エラー 2036
        Debug.Print BitShiftStrBin("1011", 2, 3, LOGICAL_SHIFT)                                 'エラー 2015
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, 3)                                    'エラー 2015
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, -1)                    'エラー 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    フォルダ選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトフォルダパス（省略可）
' = 戻値                String        フォルダ選択結果
' = 覚書    ・キャンセルが押下されると、空文字を返却する。
' =         ・存在しないフォルダを指定すると、空文字を返却する。
' ==================================================================
Public Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFolderSelectDialog = ""
    Else
        Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim sOutputDirPath As String
        sOutputDirPath = fdDialog.SelectedItems.Item(1)
        If objFSO.FolderExists(sOutputDirPath) Then
            ShowFolderSelectDialog = sOutputDirPath
        Else
            ShowFolderSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        MsgBox ShowFolderSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") _
                )
    End Sub

' ==================================================================
' = 概要    ファイル（単一）選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトファイルパス（省略可）
' = 引数    sFilters　  String  [in]  選択時のフィルタ（省略可）(※)
' = 戻値                String        ファイル選択結果
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    - 拡張子が複数ある場合は、";"で区切る
' =                    - ファイル種別と拡張子は"/"で区切る
' =                    - フィルタが複数ある場合、","で区切る
' =         ・sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' =         ・ダイアログでキャンセルが押下された場合、空文字を返却する。
' =         ・存在しないファイルは選択できない。
' ==================================================================
Public Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイルを選択してください"
    fdDialog.AllowMultiSelect = False
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFileSelectDialog = ""
    Else
        ShowFileSelectDialog = fdDialog.SelectedItems(1)
    End If
     
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFileSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        sFilters = ""
        
        MsgBox ShowFileSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
    '    MsgBox ShowFileSelectDialog( _
    '                objWshShell.SpecialFolders("Desktop") & "\test.txt" _
    '            )
    End Sub

' ==================================================================
' = 概要    ファイル（複数）選択ダイアログを表示する
' = 引数    asSelectedFiles String()    [out] 選択されたファイルパス一覧
' = 引数    sInitPath       String      [in]  デフォルトファイルパス（省略可）
' = 引数    sFilters        String      [in]  選択時のフィルタ（省略可）(※)
' = 戻値    なし
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    - 拡張子が複数ある場合は、";"で区切る
' =                    - ファイル種別と拡張子は"/"で区切る
' =                    - フィルタが複数ある場合、","で区切る
' =         ・sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' =         ・ダイアログでキャンセルが押下された場合、空文字を返却する。
' =         ・存在しないファイルは選択できない。
' ==================================================================
Public Function ShowFilesSelectDialog( _
    ByRef asSelectedFiles() As String, _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
)
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイルを選択してください（複数可）"
    fdDialog.AllowMultiSelect = True
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ReDim Preserve asSelectedFiles(0)
        asSelectedFiles(0) = ""
    Else
        Dim lSelNum As Long
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asSelectedFiles(lSelNum - 1)
        Dim lSelIdx As Long
        For lSelIdx = 0 To lSelNum - 1
            asSelectedFiles(lSelIdx) = fdDialog.SelectedItems(lSelIdx + 1)
        Next lSelIdx
    End If
     
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFilesSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        sFilters = "全てのファイル/*.*,画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        
        Dim asSelectedFiles() As String
        Call ShowFilesSelectDialog( _
                    asSelectedFiles, _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
        Dim sBuf As String
        sBuf = ""
        Dim lSelIdx As Long
        For lSelIdx = 0 To UBound(asSelectedFiles)
            sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
        Next lSelIdx
        MsgBox sBuf
    End Sub

' ==================================================================
' = 概要    セル範囲（Range型）を文字列配列（String配列型）に変換する。
' =         主にセル範囲をテキストファイルに出力する時に使用する。
' = 引数    rCellsRange             Range   [in]  対象のセル範囲
' = 引数    asLine()                String  [out] 文字列返還後のセル範囲
' = 引数    bIsInvisibleCellIgnore  String  [in]  非表示セル無視実行可否
' = 引数    sDelimiter              String  [in]  区切り文字
' = 戻値    なし
' = 覚書    列が隣り合ったセル同士は指定された区切り文字で区切られる
' ==================================================================
Private Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIsInvisibleCellIgnore As Boolean, _
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
            If bIsInvisibleCellIgnore = True Then
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

Public Sub シート毎の名前定義を削除()
    Debug.Print ThisWorkbook.Names.Count
    Dim i As Long
    Dim sName As String
    For i = ThisWorkbook.Names.Count To 1 Step -1
        sName = ThisWorkbook.Names.Item(i).Name
        'ThisWorkbook.Names.Item(i).Delete
        If Left(sName, 1) = "'" Then
            Debug.Print sName
            ThisWorkbook.Names.Item(i).Delete
        Else
            'Do Nothing
        End If
    Next i
End Sub

'********************************************************************************
'* 内部関数定義
'********************************************************************************
'Mod 演算子は 2,147,483,647 より大きい数字はオーバーフローする。
'本関数は上記以上の数値を扱うことができる。
Private Function ModEx( _
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

'32bit用
Private Function Dec2Hex( _
    ByVal cInDecVal As Currency _
) As String
    Dim cInDecValHi As Currency
    Dim cInDecValLo As Currency
    Dim sOutHexValHi As String
    Dim sOutHexValLo As String
    Dim sHexVal As String
    If cInDecVal < -2147483648# Or cInDecVal > 4294967295# Then
        Dec2Hex = "error"
        Exit Function
    End If
    
    If cInDecVal >= 0 Then
        'Do Nothing
    Else
        cInDecVal = cInDecVal + 4294967296#
    End If
    cInDecVal = Int(ModEx(cInDecVal, 4294967296#))
    
    cInDecValHi = Int(cInDecVal / 65536)
    cInDecValLo = Int(ModEx(cInDecVal, 65536))
    sOutHexValHi = UCase(String(4 - Len(Hex(cInDecValHi)), "0") & Hex(cInDecValHi))
    sOutHexValLo = UCase(String(4 - Len(Hex(cInDecValLo)), "0") & Hex(cInDecValLo))
    Dec2Hex = sOutHexValHi & sOutHexValLo
End Function
    Private Sub Test_Dec2Hex()
        Debug.Print "*** test start! ***"
        Debug.Print Dec2Hex(0)              '00000000
        Debug.Print Dec2Hex(1)              '00000001
        Debug.Print Dec2Hex(2)              '00000002
        Debug.Print Dec2Hex(10)             '0000000A
        Debug.Print Dec2Hex(15)             '0000000F
        Debug.Print Dec2Hex(16)             '00000010
        Debug.Print Dec2Hex(4294967296#)    'error
        Debug.Print Dec2Hex(4294967295#)    'FFFFFFFF
        Debug.Print Dec2Hex(4294967294#)    'FFFFFFFE
        Debug.Print Dec2Hex(2147483648#)    '80000000
        Debug.Print Dec2Hex(2147483647)     '7FFFFFFF
        Debug.Print Dec2Hex(2147483646)     '7FFFFFFE
        Debug.Print Dec2Hex(65536)          '00010000
        Debug.Print Dec2Hex(65535)          '0000FFFF
        Debug.Print Dec2Hex(65534)          '0000FFFE
        Debug.Print Dec2Hex(2)              '00000002
        Debug.Print Dec2Hex(1)              '00000001
        Debug.Print Dec2Hex(0)              '00000000
        Debug.Print Dec2Hex(-1)             'FFFFFFFF
        Debug.Print Dec2Hex(-2)             'FFFFFFFE
        Debug.Print Dec2Hex(-65534)         'FFFF0002
        Debug.Print Dec2Hex(-65535)         'FFFF0001
        Debug.Print Dec2Hex(-65536)         'FFFF0000
        Debug.Print Dec2Hex(-2147483647)    '80000001
        Debug.Print Dec2Hex(-2147483648#)   '80000000
        Debug.Print Dec2Hex(-2147483649#)   'error
        Debug.Print "*** test finished! ***"
    End Sub

'32bit用
Private Function Hex2Dec( _
    ByVal sInHexVal As String, _
    ByVal bIsSignEnable As Boolean _
) As Variant
    If Len(sInHexVal) <> 8 Then
        Hex2Dec = "error"
        Exit Function
    End If
    Dim cInDecValHi As Currency
    Dim cInDecValLo As Currency
    Dim cOutDecVal As Currency
    On Error Resume Next
    cInDecValHi = CCur("&H" & Left$(sInHexVal, 4)) * 65536
    cInDecValLo = CCur("&H" & Right$(sInHexVal, 4))
    cOutDecVal = cInDecValHi + cInDecValLo
    If Err.Number <> 0 Then
        Hex2Dec = "error"
        Err.Clear
    Else
        If bIsSignEnable = True Then
            If cOutDecVal > 2147483647 Then
                Hex2Dec = cOutDecVal - 4294967296#
            Else
                Hex2Dec = cOutDecVal
            End If
        Else
            Hex2Dec = cOutDecVal
        End If
    End If
    On Error GoTo 0
End Function
    Private Sub Test_Hex2Dec()
        Debug.Print "*** test start! ***"
        Debug.Print Hex2Dec("00000000", False) '0
        Debug.Print Hex2Dec("00000001", False) '1
        Debug.Print Hex2Dec("00000009", False) '9
        Debug.Print Hex2Dec("0000000A", False) '10
        Debug.Print "<<sign>>"
        Debug.Print Hex2Dec("7FFFFFFF", True)  '2147483647
        Debug.Print Hex2Dec("7FFFFFFE", True)  '2147483646
        Debug.Print Hex2Dec("00000002", True)  '2
        Debug.Print Hex2Dec("00000001", True)  '1
        Debug.Print Hex2Dec("00000000", True)  '0
        Debug.Print Hex2Dec("FFFFFFFF", True)  '-1
        Debug.Print Hex2Dec("FFFFFFFE", True)  '-2
        Debug.Print Hex2Dec("80000001", True)  '-2147483647
        Debug.Print Hex2Dec("80000000", True)  '-2147483648
        Debug.Print "<<unsign>>"
        Debug.Print Hex2Dec("FFFFFFFF", False) '4294967295
        Debug.Print Hex2Dec("FFFFFFFE", False) '4294967294
        Debug.Print Hex2Dec("80000001", False) '2147483649
        Debug.Print Hex2Dec("80000000", False) '2147483648
        Debug.Print Hex2Dec("00000001", False) '1
        Debug.Print Hex2Dec("00000000", False) '0
        Debug.Print Hex2Dec("0000000", False)  'error
        Debug.Print Hex2Dec("000000000", False) 'error
        Debug.Print Hex2Dec("8000000K", False) 'error
        Debug.Print Hex2Dec("80 00001", False) 'error
        Debug.Print "*** test finished! ***"
    End Sub

'指定範囲以外の値を指定すると文字列 "error" を返却する。
Private Function Hex2Bin( _
    ByVal sHexVal As String _
) As String
    Dim sOutBinVal As String
    Dim sTmpBinVal As String
    Dim lIdx As Long
    Dim sChar As String
    If sHexVal = "" Then
        sOutBinVal = ""
    Else
        sOutBinVal = ""
        For lIdx = 1 To Len(sHexVal)
            sChar = Mid$(sHexVal, lIdx, 1)
            sTmpBinVal = Hex2BinMap(sChar)
            If sTmpBinVal = "error" Then
                sOutBinVal = sTmpBinVal
                Exit For
            Else
                sOutBinVal = sOutBinVal & sTmpBinVal
            End If
        Next lIdx
    End If
    Hex2Bin = sOutBinVal
End Function
    Private Sub Test_Hex2Bin()
        Debug.Print "*** test start! ***"
        Debug.Print Hex2Bin("0123")      '0000000100100011
        Debug.Print Hex2Bin("4567")      '0100010101100111
        Debug.Print Hex2Bin("89AB")      '1000100110101011
        Debug.Print Hex2Bin("CDEF")      '1100110111101111
        Debug.Print Hex2Bin("cdef")      '1100110111101111
        Debug.Print Hex2Bin("c")         '1100
        Debug.Print Hex2Bin("01234567")  '00000001001000110100010101100111
        Debug.Print Hex2Bin("")          '
        Debug.Print Hex2Bin("ab ")       'error
        Debug.Print Hex2Bin(":cd")       'error
        Debug.Print "*** test finished! ***"
    End Sub

'指定範囲以外の値を指定すると文字列 "error" を返却する。
Private Function Bin2Hex( _
    ByVal sBinVal As String, _
    ByVal bIsUcase As Boolean _
) As String
    Dim sExtBinStr As String
    Dim sTmpHexVal As String
    Dim sOutHexVal As String
    Dim lIdx As Long
    If sBinVal = "" Then
        sOutHexVal = ""
    Else
        If Len(sBinVal) Mod 4 = 0 Then
            For lIdx = 1 To Len(sBinVal) Step 4
                sExtBinStr = Mid$(sBinVal, lIdx, 4)
                sTmpHexVal = Bin2HexMap(sExtBinStr, bIsUcase)
                If sTmpHexVal = "error" Then
                    sOutHexVal = "error"
                    Exit For
                Else
                    sOutHexVal = sOutHexVal & sTmpHexVal
                End If
            Next lIdx
        Else
            sOutHexVal = "error"
        End If
    End If
    Bin2Hex = sOutHexVal
End Function
    Private Sub Test_Bin2Hex()
        Debug.Print "*** test start! ***"
        Debug.Print Bin2Hex("0000000100100011", True)                   '0123
        Debug.Print Bin2Hex("0100010101100111", True)                   '4567
        Debug.Print Bin2Hex("1000100110101011", True)                   '89AB
        Debug.Print Bin2Hex("1100110111101111", True)                   'CDEF
        Debug.Print Bin2Hex("1100110111101111", False)                  'cdef
        Debug.Print Bin2Hex("1100", False)                              'c
        Debug.Print Bin2Hex("00000001001000110100010101100111", True)   '01234567
        Debug.Print Bin2Hex("", True)                                   '
        Debug.Print Bin2Hex("110011011110111", False)                   'error
        Debug.Print Bin2Hex("010 ", True)                               'error
        Debug.Print Bin2Hex(":011", True)                               'error
        Debug.Print "*** test finished! ***"
    End Sub

'指定範囲以外の値を指定すると文字列 "error" を返却する。
Private Function Hex2BinMap( _
    ByVal sHexVal As String _
) As String
    Select Case UCase(sHexVal)
        Case "0":  Hex2BinMap = "0000"
        Case "1":  Hex2BinMap = "0001"
        Case "2":  Hex2BinMap = "0010"
        Case "3":  Hex2BinMap = "0011"
        Case "4":  Hex2BinMap = "0100"
        Case "5":  Hex2BinMap = "0101"
        Case "6":  Hex2BinMap = "0110"
        Case "7":  Hex2BinMap = "0111"
        Case "8":  Hex2BinMap = "1000"
        Case "9":  Hex2BinMap = "1001"
        Case "A":  Hex2BinMap = "1010"
        Case "B":  Hex2BinMap = "1011"
        Case "C":  Hex2BinMap = "1100"
        Case "D":  Hex2BinMap = "1101"
        Case "E":  Hex2BinMap = "1110"
        Case "F":  Hex2BinMap = "1111"
        Case Else: Hex2BinMap = "error"
    End Select
End Function

'指定範囲以外の値を指定すると文字列 "error" を返却する。
Private Function Bin2HexMap( _
    ByVal sBinVal As String, _
    ByVal bIsUcase As Boolean _
) As String
    If bIsUcase = True Then
        Select Case sBinVal
            Case "0000": Bin2HexMap = "0"
            Case "0001": Bin2HexMap = "1"
            Case "0010": Bin2HexMap = "2"
            Case "0011": Bin2HexMap = "3"
            Case "0100": Bin2HexMap = "4"
            Case "0101": Bin2HexMap = "5"
            Case "0110": Bin2HexMap = "6"
            Case "0111": Bin2HexMap = "7"
            Case "1000": Bin2HexMap = "8"
            Case "1001": Bin2HexMap = "9"
            Case "1010": Bin2HexMap = "A"
            Case "1011": Bin2HexMap = "B"
            Case "1100": Bin2HexMap = "C"
            Case "1101": Bin2HexMap = "D"
            Case "1110": Bin2HexMap = "E"
            Case "1111": Bin2HexMap = "F"
            Case Else:   Bin2HexMap = "error"
        End Select
    Else
        Select Case sBinVal
            Case "0000": Bin2HexMap = "0"
            Case "0001": Bin2HexMap = "1"
            Case "0010": Bin2HexMap = "2"
            Case "0011": Bin2HexMap = "3"
            Case "0100": Bin2HexMap = "4"
            Case "0101": Bin2HexMap = "5"
            Case "0110": Bin2HexMap = "6"
            Case "0111": Bin2HexMap = "7"
            Case "1000": Bin2HexMap = "8"
            Case "1001": Bin2HexMap = "9"
            Case "1010": Bin2HexMap = "a"
            Case "1011": Bin2HexMap = "b"
            Case "1100": Bin2HexMap = "c"
            Case "1101": Bin2HexMap = "d"
            Case "1110": Bin2HexMap = "e"
            Case "1111": Bin2HexMap = "f"
            Case Else:   Bin2HexMap = "error"
        End Select
    End If
End Function

'論理ビットシフト（文字列版）
Private Function BitShiftLogStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal lInBitLen As Long _
) As String
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
    Debug.Assert lInBitLen >= 0
    
    'ビットシフト
    Dim sTmpBinVal As String
    Select Case eInDirection
        Case RIGHT_SHIFT
            If Len(sInBinVal) > lInShiftNum Then
                sTmpBinVal = Left$(sInBinVal, Len(sInBinVal) - lInShiftNum)
            Else
                sTmpBinVal = "0"
            End If
        Case LEFT_SHIFT
            sTmpBinVal = sInBinVal & String(lInShiftNum, "0")
        Case Else
            Debug.Assert 0
    End Select
    
    'ビット位置合わせ
    If lInBitLen = 0 Then
        BitShiftLogStrBin = sTmpBinVal
    Else
        If lInBitLen > Len(sTmpBinVal) Then
            BitShiftLogStrBin = String(lInBitLen - Len(sTmpBinVal), "0") & sTmpBinVal
        ElseIf lInBitLen < Len(sTmpBinVal) Then
            BitShiftLogStrBin = Right$(sTmpBinVal, lInBitLen)
        Else
            BitShiftLogStrBin = sTmpBinVal
        End If
    End If
End Function
    Private Sub Test_BitShiftLogStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("0", 2, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("1", 0, LEFT_SHIFT, 8)       '00000001
        Debug.Print BitShiftLogStrBin("1", 2, LEFT_SHIFT, 8)       '00000100
        Debug.Print BitShiftLogStrBin("1", 7, LEFT_SHIFT, 8)       '10000000
        Debug.Print BitShiftLogStrBin("1", 8, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("1", 0, RIGHT_SHIFT, 8)      '00000001
        Debug.Print BitShiftLogStrBin("1", 1, RIGHT_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1", 2, RIGHT_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1011", 0, LEFT_SHIFT, 0)    '1011
        Debug.Print BitShiftLogStrBin("1011", 1, LEFT_SHIFT, 0)    '10110
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 0)    '101100
        Debug.Print BitShiftLogStrBin("1011", 0, RIGHT_SHIFT, 0)   '1011
        Debug.Print BitShiftLogStrBin("1011", 1, RIGHT_SHIFT, 0)   '101
        Debug.Print BitShiftLogStrBin("1011", 2, RIGHT_SHIFT, 0)   '10
        Debug.Print BitShiftLogStrBin("1011", 3, RIGHT_SHIFT, 0)   '1
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 0)   '0
        Debug.Print BitShiftLogStrBin("1011", 5, RIGHT_SHIFT, 0)   '0
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 2)    '00
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 3)    '100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 4)    '1100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 10)   '0000101100
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 8)   '00000000
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 8)   '00000000
       'Debug.Print BitShiftLogStrBin("", 2, LEFT_SHIFT,  10)       'プログラム停止
       'Debug.Print BitShiftLogStrBin("101A", 1, LEFT_SHIFT,  10)   'プログラム停止
       'Debug.Print BitShiftLogStrBin("1011", -1, LEFT_SHIFT,  10)  'プログラム停止
       'Debug.Print BitShiftLogStrBin("1011", 1, 5,  10)            'プログラム停止
        Debug.Print "*** test finished! ***"
    End Sub

'算術ビットシフト（文字列版）
Private Function BitShiftAriStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal lInBitLen As Long, _
    ByVal bIsSaveSignBit As Boolean, _
    ByVal bIsExecAutoAlign As Boolean _
) As String
    ' <<bIsSaveSignBit>>
    '   出力桁数が入力値（文字列）の長さよりも小さい場合に、
    '   符号ビットを保持するか、無視して切り捨てるかを選択する。
    '     True  : 符号ビットを保持する
    '               ex1) "AB"(0b10101011) を出力桁数1として右1算術(符号ビット保持)シフト
    '                 ⇒ "D"(0b1101)
    '     False : 符号ビットを切り捨てる
    '               ex2) "AB"(0b10101011) を出力桁数1として右1算術(符号ビット切捨)シフト
    '                 ⇒ "5"(0b0101)
    ' <<bIsExecAutoAlign>>
    '   出力結果を8ビット境界に揃えるかどうかを選択する。
    '     True  : 揃える
    '               ex1) 10101011 を右1ビットシフト
    '                 ⇒ 11010101
    '               ex2) 10101011 を左1ビットシフト
    '                 ⇒ 1111111101010110
    '     False : 揃えない
    '               ex1) 10101011 を右1ビットシフト
    '                 ⇒  1010101
    '               ex2) 10101011 を左1ビットシフト
    '                 ⇒ 101010110
    
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
    Debug.Assert lInBitLen >= 0
    If bIsExecAutoAlign = True Then
        Debug.Assert Len(sInBinVal) = 8
        Debug.Assert lInBitLen Mod 8 = 0
    Else
        'Do Nothing
    End If
    
    'ビットシフト
    Dim sTmpBinVal As String
    Dim sOutLogicBit As String
    Dim sInSignBit As String
    Dim sInLogicBit As String
    sInSignBit = Left$(sInBinVal, 1)
    sInLogicBit = Mid$(sInBinVal, 2, Len(sInBinVal))
    Select Case eInDirection
        Case RIGHT_SHIFT
            If Len(sInLogicBit) > lInShiftNum Then
                sOutLogicBit = Left$(sInLogicBit, Len(sInLogicBit) - lInShiftNum)
            Else
                sOutLogicBit = ""
            End If
        Case LEFT_SHIFT
            sOutLogicBit = sInLogicBit & String(lInShiftNum, "0")
        Case Else
            Debug.Assert 0
    End Select
    sTmpBinVal = sInSignBit & sOutLogicBit
    
    'ビット位置合わせ
    If lInBitLen = 0 Then
        If bIsExecAutoAlign = True Then
            Dim sPadBit As String
            If ((Len(sOutLogicBit) + 1) Mod 8) = 0 Then
                sPadBit = ""
            Else
                sPadBit = String(8 - ((Len(sOutLogicBit) + 1) Mod 8), sInSignBit)
            End If
            BitShiftAriStrBin = sPadBit & sInSignBit & sOutLogicBit
        Else
            BitShiftAriStrBin = sTmpBinVal
        End If
    Else
        If lInBitLen > Len(sTmpBinVal) Then
            BitShiftAriStrBin = String(lInBitLen - Len(sTmpBinVal), sInSignBit) & sTmpBinVal
        ElseIf lInBitLen < Len(sTmpBinVal) Then
            If bIsSaveSignBit = True Then
                BitShiftAriStrBin = sInSignBit & Right$(sOutLogicBit, lInBitLen - 1)
            Else
                BitShiftAriStrBin = Right$(sTmpBinVal, lInBitLen)
            End If
        Else
            BitShiftAriStrBin = sTmpBinVal
        End If
    End If
End Function
    Private Sub Test_BitShiftAriStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print "<<test 001-01>>"                                                   '<<test 001-01>>
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 0, True, False)       '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, 0, True, False)       '0100101
        Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, 0, True, False)       '010010
        Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, 0, True, False)       '01001
        Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, 0, True, False)       '0100
        Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, 0, True, False)       '010
        Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, 0, True, False)       '01
        Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 10, True, False)      '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 5, True, False)       '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 4, True, False)       '0011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 3, True, False)       '011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 2, True, False)       '01
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 1, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 0, True, False)        '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, 0, True, False)        '010010110
        Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, 0, True, False)        '0100101100
        Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, 0, True, False)        '0100101100000
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 10, True, False)       '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 5, True, False)        '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 4, True, False)        '0011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 3, True, False)        '011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 2, True, False)        '01
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 1, True, False)        '0
        Debug.Print "<<test 001-02>>"                                                   '<<test 001-02>>
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 0, False, False)      '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, 0, False, False)      '0100101
        Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, 0, False, False)      '010010
        Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, 0, False, False)      '01001
        Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, 0, False, False)      '0100
        Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, 0, False, False)      '010
        Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, 0, False, False)      '01
        Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 10, False, False)     '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 5, False, False)      '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 4, False, False)      '1011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 3, False, False)      '011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 2, False, False)      '11
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 1, False, False)      '1
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 0, False, False)       '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, 0, False, False)       '010010110
        Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, 0, False, False)       '0100101100
        Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, 0, False, False)       '0100101100000
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 10, False, False)      '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 5, False, False)       '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 4, False, False)       '1011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 3, False, False)       '011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 2, False, False)       '11
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 1, False, False)       '1
        Debug.Print "<<test 001-03>>"                                                   '<<test 001-03>>
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 0, True, False)       '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, 0, True, False)       '1000101
        Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, 0, True, False)       '100010
        Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, 0, True, False)       '10001
        Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, 0, True, False)       '1000
        Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, 0, True, False)       '100
        Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, 0, True, False)       '10
        Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 10, True, False)      '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 5, True, False)       '11011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 4, True, False)       '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 3, True, False)       '111
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 2, True, False)       '11
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 1, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 0, True, False)        '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, 0, True, False)        '100010110
        Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, 0, True, False)        '1000101100
        Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, 0, True, False)        '1000101100000
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 10, True, False)       '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 5, True, False)        '11011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 4, True, False)        '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 3, True, False)        '111
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 2, True, False)        '11
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 1, True, False)        '1
        Debug.Print "<<test 001-04>>"                                                   '<<test 001-04>>
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 0, False, False)      '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, 0, False, False)      '1000101
        Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, 0, False, False)      '100010
        Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, 0, False, False)      '10001
        Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, 0, False, False)      '1000
        Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, 0, False, False)      '100
        Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, 0, False, False)      '10
        Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 10, False, False)     '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 5, False, False)      '01011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 4, False, False)      '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 3, False, False)      '011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 2, False, False)      '11
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 1, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 0, False, False)       '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, 0, False, False)       '100010110
        Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, 0, False, False)       '1000101100
        Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, 0, False, False)       '1000101100000
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 10, False, False)      '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 5, False, False)       '01011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 4, False, False)       '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 3, False, False)       '011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 2, False, False)       '11
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 1, False, False)       '1
    '   Debug.Print "<<test 001-05>>"                                                   '<<test 001-05>>
    '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, 10, True, False)                'プログラム停止
    '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, 10, True, False)               'プログラム停止
    '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, 10, True, False)      'プログラム停止
    '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, 10, True, False)       'プログラム停止
    '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, -1, True, False)       'プログラム停止
        Debug.Print "<<test 002-01>>"                                                   '<<test 002-01>>"
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 0, True, True)        '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, 0, True, True)        '00010101
        Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, 0, True, True)        '00001010
        Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, 0, True, True)        '00000101
        Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, 0, True, True)        '00000010
        Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, 0, True, True)        '00000001
        Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, 0, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, 0, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 16, True, True)       '0000000000101011
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 8, True, True)        '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, 8, True, True)        '00010101
        Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, 8, True, True)        '00001010
        Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, 8, True, True)        '00000101
        Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, 8, True, True)        '00000010
        Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, 8, True, True)        '00000001
        Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, 8, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, 8, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 0, True, True)         '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, 0, True, True)         '0000000001010110
        Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, 0, True, True)         '0000000010101100
        Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, 0, True, True)         '0000010101100000
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 16, True, True)        '0000000000101011
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 8, True, True)         '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, 8, True, True)         '01010110
        Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, 8, True, True)         '00101100
        Debug.Print BitShiftAriStrBin("00101011", 3, LEFT_SHIFT, 8, True, True)         '01011000
        Debug.Print BitShiftAriStrBin("00101011", 4, LEFT_SHIFT, 8, True, True)         '00110000
        Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, 8, True, True)         '01100000
        Debug.Print BitShiftAriStrBin("00101011", 6, LEFT_SHIFT, 8, True, True)         '01000000
        Debug.Print BitShiftAriStrBin("00101011", 7, LEFT_SHIFT, 8, True, True)         '00000000
        Debug.Print "<<test 002-02>>"                                                   '<<test 002-02>>
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 0, True, True)        '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, 0, True, True)        '11010101
        Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, 0, True, True)        '11101010
        Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, 0, True, True)        '11110101
        Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, 0, True, True)        '11111010
        Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, 0, True, True)        '11111101
        Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, 0, True, True)        '11111110
        Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, 0, True, True)        '11111111
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 16, True, True)       '1111111110101011
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 8, True, True)        '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, 8, True, True)        '11010101
        Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, 8, True, True)        '11101010
        Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, 8, True, True)        '11110101
        Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, 8, True, True)        '11111010
        Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, 8, True, True)        '11111101
        Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, 8, True, True)        '11111110
        Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, 8, True, True)        '11111111
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 0, True, True)         '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, 0, True, True)         '1111111101010110
        Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, 0, True, True)         '1111111010101100
        Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, 0, True, True)         '1111010101100000
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 16, True, True)        '1111111110101011
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 8, True, True)         '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, 8, True, True)         '11010110
        Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, 8, True, True)         '10101100
        Debug.Print BitShiftAriStrBin("10101011", 3, LEFT_SHIFT, 8, True, True)         '11011000
        Debug.Print BitShiftAriStrBin("10101011", 4, LEFT_SHIFT, 8, True, True)         '10110000
        Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, 8, True, True)         '11100000
        Debug.Print BitShiftAriStrBin("10101011", 6, LEFT_SHIFT, 8, True, True)         '11000000
        Debug.Print BitShiftAriStrBin("10101011", 7, LEFT_SHIFT, 8, True, True)         '10000000
        Debug.Print BitShiftAriStrBin("10101011", 8, LEFT_SHIFT, 8, True, True)         '10000000
    '   Debug.Print "<<test 002-03>>"                                                   '<<test 002-03>>
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 8, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 5, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 4, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 3, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 2, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 1, True, True)         'プログラム停止
    '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, 10, True, True)                 'プログラム停止
    '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, 10, True, True)                'プログラム停止
    '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, 10, True, True)       'プログラム停止
    '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, 10, True, True)        'プログラム停止
    '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, -1, True, True)        'プログラム停止
        Debug.Print "*** test finished! ***"
    End Sub

'ShowFileSelectDialog() と ShowFilesSelectDialog() 用の関数
'ダイアログのフィルタを追加する。指定方法は以下。
'  ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
'      ・拡張子が複数ある場合は、";"で区切る
'      ・ファイル種別と拡張子は"/"で区切る
'      ・フィルタが複数ある場合、","で区切る
'sFilters が空文字の場合、フィルタをクリアする。
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                            """/"" は一つだけ指定してください" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "処理を中断します。"
                        End
                    End If
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "処理を中断します。"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        """/"" は一つだけ指定してください" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "処理を中断します。"
                    End
                End If
            Else
                MsgBox _
                    "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                    "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                    "  " & sFilters
                MsgBox "処理を中断します。"
                End
            End If
        End If
    End If
End Function

