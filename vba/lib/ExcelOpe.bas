Attribute VB_Name = "ExcelOpe"
Option Explicit

' excel operation library v1.0

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
' = 戻値                  String        結合後の文字列
' = 覚書    なし
' ==================================================================
Public Function ConcStr( _
    ByRef rConcRange As Range, _
    Optional ByVal sDlmtr As String _
) As Variant
    Dim rConcRangeCnt As Range
    Dim sConcTxtBuf As String
 
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
End Function
 
' ==================================================================
' = 概要    文字列を分割し、指定した要素の文字列を返却する
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 引数    iExtIndex   String  [in]  抽出する要素 ( 0 origin )
' = 戻値                String        抽出文字列
' = 覚書    iExtIndex が要素を超える場合、空文字列を返却する
' ==================================================================
Public Function SplitStr( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As String
    Dim vSplitStr As Variant
 
    ' 文字列分割
    vSplitStr = Split(sStr, sDlmtr)
 
    If iExtIndex > UBound(vSplitStr) Then
        SplitStr = ""
    Else
        SplitStr = vSplitStr(iExtIndex)
    End If
End Function

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
 
' ==================================================================
' = 概要    フォントカラーを返却する
' = 引数    rRange   Range     [in]  セル
' = 戻値             Variant         フォントカラー
' = 覚書    なし
' ==================================================================
Public Function GetFontColor( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetFontColor = rRange.Font.Color
    Else
        GetFontColor = CVErr(xlErrRef)  'エラー値
    End If
End Function
 
' ==================================================================
' = 概要    背景色を返却する
' = 引数    rRange   Range     [in]  セル
' = 戻値             Variant         背景色
' = 覚書    なし
' ==================================================================
Public Function GetInteriorColor( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetInteriorColor = rRange.Interior.Color
    Else
        GetInteriorColor = CVErr(xlErrRef)  'エラー値
    End If
End Function

' ==================================================================
' = 概要    ビット AND 演算を行う
' = 引数    cInVar1   Currency   [in]  入力値 左項（10進数数値）
' = 引数    cInVar2   Currency   [in]  入力値 右項（10進数数値）
' = 戻値              Variant          演算結果（10進数数値）
' = 覚書    なし
' ==================================================================
Public Function BitAnd( _
    ByVal cInVar1 As Currency, _
    ByVal cInVar2 As Currency _
) As Variant
    If (cInVar1 > 2147483647# Or cInVar2 > 2147483647#) Then
        BitAnd = CVErr(xlErrNum)  'エラー値
    Else
        BitAnd = cInVar1 And cInVar2
    End If
End Function
 
' ==================================================================
' = 概要    ビット OR 演算を行う
' = 引数    cInVar1   Currency   [in]  入力値 左項（10進数数値）
' = 引数    cInVar2   Currency   [in]  入力値 右項（10進数数値）
' = 戻値              Variant          演算結果（10進数数値）
' = 覚書    なし
' ==================================================================
Public Function BitOr( _
    ByVal cInVar1 As Currency, _
    ByVal cInVar2 As Currency _
) As Variant
    Dim sHexVal As String
    If (cInVar1 > 2147483647# Or cInVar2 > 2147483647#) Then
        BitOr = CVErr(xlErrNum)  'エラー値
    Else
        BitOr = cInVar1 Or cInVar2
    End If
End Function
 
' ==================================================================
' = 概要    論理シフトを行う。
' = 引数    cDecVal     Currency  [in]  入力値（10進数数値）
' = 引数    lShiftNum   Long      [in]  シフトビット数
' = 引数    eDirection  Enum      [in]  シフト方向（右:0 左:1）
' = 引数    eShiftType  Enum      [in]  シフト種別（右:論理 左:算術）
' = 戻値                Variant         シフト結果（10進数数値）
' = 覚書    なし
' ==================================================================
Public Function BitShift( _
    ByVal cDecVal As Currency, _
    ByVal lShiftNum As Long, _
    ByVal eDirection As E_SHIFT_DIRECTiON, _
    ByVal eShiftType As E_SHIFT_TYPE _
) As Variant
    Dim sHexVal As String
    Dim cDecValHi As Currency
    Dim cDecValLo As Currency
    Dim sBinVal As String
    Dim cRetVal As Currency
 
    If cDecVal > 4294967295# Or _
       (lShiftNum < 0) Or _
       (eDirection <> RIGHT_SHIFT And eDirection <> LEFT_SHIFT) Then
        BitShift = CVErr(xlErrNum)  'エラー値
    Else
        If eShiftType = LOGICAL_SHIFT Then
            'Dec⇒Hex
            cDecValHi = Int(cDecVal / 2 ^ 16)
            cDecValLo = cDecVal - (cDecValHi * 2 ^ 16)
            sHexVal = UCase(String(4 - Len(Hex(cDecValHi)), "0") & Hex(cDecValHi)) & _
                      UCase(String(4 - Len(Hex(cDecValLo)), "0") & Hex(cDecValLo))
            'Hex⇒Bin
            sBinVal = Hex2Bin(sHexVal)
            'Shift
            sBinVal = BitLogShiftBin(sBinVal, lShiftNum, eDirection)
            'Bin⇒Hex
            sHexVal = Bin2Hex(sBinVal)
            'Hex⇒Dec
            cDecValHi = CCur("&H" & Left$(sHexVal, 4)) * 2 ^ 16
            cDecValLo = CCur("&H" & Right$(sHexVal, 4))
            BitShift = cDecValHi + cDecValLo
        Else
            BitShift = CVErr(xlErrNum) '算術シフトは非対応
        End If
    End If
End Function
 
'********************************************************************************
'* 内部関数定義
'********************************************************************************
Private Function Hex2Bin( _
    ByVal sHexVal As String _
) As String
    Dim sBinVal As String
    Debug.Assert Len(sHexVal) = 8
    Do
        sBinVal = sBinVal & Hex2BinMap(Left$(sHexVal, 1))
        sHexVal = Mid$(sHexVal, 2)
    Loop While sHexVal <> ""
    Hex2Bin = sBinVal
End Function
 
Private Function Bin2Hex( _
    ByVal sBinVal As String _
) As String
    Dim sHexVal As String
    Debug.Assert Len(sBinVal) = 32
    Do
        sHexVal = sHexVal & Bin2HexMap(Left$(sBinVal, 4))
        sBinVal = Mid$(sBinVal, 5)
    Loop While sBinVal <> ""
    Bin2Hex = sHexVal
End Function
 
Private Function BitLogShiftBin( _
    ByVal sBinVal As String, _
    ByVal lShiftNum As Long, _
    ByVal eDirection As E_SHIFT_DIRECTiON _
)
    Debug.Assert Len(sBinVal) = 32
    Debug.Assert lShiftNum >= 0
    If lShiftNum < 32 Then
        Select Case eDirection
            Case RIGHT_SHIFT
                BitLogShiftBin = String(lShiftNum, "0") & Left$(sBinVal, Len(sBinVal) - lShiftNum)
            Case LEFT_SHIFT
                BitLogShiftBin = Right$(sBinVal, Len(sBinVal) - lShiftNum) & String(lShiftNum, "0")
            Case Else
                Debug.Assert False
        End Select
    Else
        BitLogShiftBin = "00000000000000000000000000000000"
    End If
End Function
 
Private Function Hex2BinMap( _
    ByVal sHexVal As String _
) As String
    Select Case sHexVal
        Case "0": Hex2BinMap = "0000"
        Case "1": Hex2BinMap = "0001"
        Case "2": Hex2BinMap = "0010"
        Case "3": Hex2BinMap = "0011"
        Case "4": Hex2BinMap = "0100"
        Case "5": Hex2BinMap = "0101"
        Case "6": Hex2BinMap = "0110"
        Case "7": Hex2BinMap = "0111"
        Case "8": Hex2BinMap = "1000"
        Case "9": Hex2BinMap = "1001"
        Case "A": Hex2BinMap = "1010"
        Case "B": Hex2BinMap = "1011"
        Case "C": Hex2BinMap = "1100"
        Case "D": Hex2BinMap = "1101"
        Case "E": Hex2BinMap = "1110"
        Case "F": Hex2BinMap = "1111"
        Case Else: Debug.Assert False
    End Select
End Function
 
Private Function Bin2HexMap( _
    ByVal sBinVal As String _
) As String
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
        Case Else: Debug.Assert False
    End Select
End Function

