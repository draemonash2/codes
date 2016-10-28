Attribute VB_Name = "Funcs"
Option Explicit

' user define functions v1.0

' ==================================================================
' =  <<関数一覧>>
' =    ・ConcStr            指定した範囲の文字列を結合する。
' =    ・SplitStr           文字列を分割し、指定した要素の文字列を返却する。
' =    ・GetStrNum          指定文字列の個数を返却する。
' =    ・RemoveTailWord     末尾区切り文字以降の文字列を除去する｡
' =    ・ExtractTailWord    末尾区切り文字以降の文字列を返却する｡
' =    ・GetDirPath         指定されたファイルパスからフォルダパスを抽出する。
' =    ・GetFileName        指定されたファイルパスからファイル名を抽出する。
' =    ・GetStrikeExist     取り消し線の有無を判定する。
' =    ・GetFontColor       フォントカラーを返却する。
' =    ・GetInteriorColor   背景色を返却する。
' =    ・BitAnd             ビット AND 演算を行う。
' =    ・BitOr              ビット OR 演算を行う。
' =    ・BitShift           論理シフトを行う。
' =    ・RegExpSearch       正規表現検索を行う。
' =    ・ConvSnakeToPascal  命名規則変換を行う（スネークケース⇒パスカルケース）
' =    ・ConvSnakeToCamel   命名規則変換を行う（スネークケース⇒キャメルケース）
' =    ・ConvCamelToSnake   命名規則変換を行う（キャメルケース⇒スネークケース）
' ==================================================================

'********************************************************************************
'* 定数定義
'********************************************************************************
Public Enum E_SHIFT_DIRECTiON
    RIGHT_SHIFT = 0
    LEFT_SHIFT
End Enum
Public Enum E_SHIFT_TYPE
    LOGICAL_SHIFT = 0
    ARITHMETIC_SHIFT '非対応
End Enum

'********************************************************************************
'* 外部関数定義
'********************************************************************************
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
' = 概要    指定文字列の個数を返却する。
' = 引数    sTrgtStr      String  [in]  検索対象文字列
' = 引数    sSrchStr      String  [in]  検索文字列
' = 戻値                  Long          文字列の個数
' = 覚書    SplitStr との組み合わせでファイル名取り出しが可能。
' =           ex) B1 = C:\codes\c\Try04.c
' =               B2 = SplitStr( B1, "\", GetStrNum( B2, "\" ) )
' =                 ⇒ Try04.c
' ==================================================================
Public Function GetStrNum( _
    ByVal sTrgtStr As String, _
    ByVal sSrchStr As String _
) As Long
    Dim vSplitStr As Variant
    
    ' 文字列分割
    vSplitStr = Split(sTrgtStr, sSrchStr)
    
    GetStrNum = UBound(vSplitStr)
End Function

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim asSplitWord() As String
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
Sub Test_ExtractTailWord()
    Debug.Print ExtractTailWord("", "\")               '
    Debug.Print ExtractTailWord("c:\a", "\")           ' a
    Debug.Print ExtractTailWord("c:\a\", "\")          '
    Debug.Print ExtractTailWord("c:\a\b", "\")         ' b
    Debug.Print ExtractTailWord("c:\a\b\", "\")        '
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "\")   ' c.txt
    Debug.Print ExtractTailWord("c:\\b\c.txt", "\")    ' c.txt
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
    Debug.Print ExtractTailWord("c:\a\\b\c.txt", "\\") ' b\c.txt
End Sub

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim sTailWord As String
    Dim lRemoveLen As Long
    
    If sStr = "" Then
        RemoveTailWord = ""
    Else
        If sDlmtr = "" Then
            RemoveTailWord = sStr
        Else
            If InStr(sStr, sDlmtr) = 0 Then
                RemoveTailWord = sStr
            Else
                sTailWord = ExtractTailWord(sStr, sDlmtr)
                lRemoveLen = Len(sDlmtr) + Len(sTailWord)
                RemoveTailWord = Left$(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function
Private Sub Test_RemoveTailWord()
    Debug.Print RemoveTailWord("", "\") '
    Debug.Print RemoveTailWord("c:\a", "\") '          c:
    Debug.Print RemoveTailWord("c:\a\", "\") '         c:\a
    Debug.Print RemoveTailWord("c:\a\b", "\") '        c:\a
    Debug.Print RemoveTailWord("c:\a\b\", "\") '       c:\a\b
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "\") '  c:\a\b
    Debug.Print RemoveTailWord("c:\\b\c.txt", "\") '   c:\\b
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "") '   c:\a\b\c.txt
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "\\") ' c:\a\b\c.txt
    Debug.Print RemoveTailWord("c:\a\\b\c.txt", "\\") 'c:\a
End Sub

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        フォルダパス
' = 覚書    なし
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String _
) As String
    GetDirPath = RemoveTailWord(sFilePath, "\")
End Function

' ==================================================================
' = 概要    指定されたファイルパスからファイル名を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイル名
' = 覚書    なし
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String _
) As String
    GetFileName = ExtractTailWord(sFilePath, "\")
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

' ==================================================================
' = 概要    正規表現検索を行う
' = 引数    sSearchPattern  String   [in]  検索パターン
' = 引数    sTargetStr      String   [in]  検索対象文字列
' = 引数    lMatchIdx       Long     [in]  検索結果インデックス（引数省略可）
' = 引数    bIsIgnoreCase   Boolean  [in]  大/小文字区別しないか（引数省略可）
' = 引数    bIsGlobal       Boolean  [in]  文字列全体を検索するか（引数省略可）
' = 戻値                    Variant        検索結果
' = 覚書    なし
' ==================================================================
Public Function RegExpSearch( _
    ByVal sSearchPattern As String, _
    ByVal sTargetStr As String, _
    Optional ByVal lMatchIdx As Long = 0, _
    Optional ByVal bIsIgnoreCase As Boolean = True, _
    Optional ByVal bIsGlobal As Boolean = True _
) As Variant
    Dim oMatchResult As Object
    Dim oRegExp As Object
    
    Set oRegExp = CreateObject("VBScript.RegExp")
    
    oRegExp.Pattern = sSearchPattern               '検索パターンを設定
    oRegExp.IgnoreCase = bIsIgnoreCase             '大文字と小文字を区別しない
    oRegExp.Global = bIsGlobal                     '文字列全体を検索
    
    Set oMatchResult = oRegExp.Execute(sTargetStr) 'パターンマッチ実行
    
    If lMatchIdx < 0 Or lMatchIdx > oMatchResult.Count - 1 Then
        RegExpSearch = CVErr(xlErrValue)  'エラー値
    Else
        RegExpSearch = oMatchResult(lMatchIdx).Value
    End If
End Function

' ==================================================================
' = 概要    命名規則変換を行う（スネークケース⇒パスカルケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                String        文字列（パスカルケース）
' = 覚書    スネークケース                          … get_input_reader
'           パスカルケース（アッパーキャメルケース）… GetInputReader
' ==================================================================
Public Function ConvSnakeToPascal( _
    ByVal sInStr As String _
) As String
    sInStr = Replace(sInStr, "_", " ")
    sInStr = StrConv(sInStr, vbProperCase)
    sInStr = Replace(sInStr, " ", "")
    ConvSnakeToPascal = sInStr
End Function

' ==================================================================
' = 概要    命名規則変換を行う（スネークケース⇒キャメルケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                String        文字列（キャメルケース）
' = 覚書    スネークケース                          … get_input_reader
'           キャメルケース（ローワーキャメルケース）… getInputReader
' ==================================================================
Public Function ConvSnakeToCamel( _
    ByVal sInStr As String _
) As String
    If sInStr = "" Then Exit Function
    
    sInStr = Replace(sInStr, "_", " ")
    sInStr = StrConv(sInStr, vbProperCase)
    sInStr = Replace(sInStr, " ", "")
    sInStr = LCase(Left$(sInStr, 1)) & _
             Mid$(sInStr, 2, Len(sInStr))
    ConvSnakeToCamel = sInStr
End Function

' ==================================================================
' = 概要    命名規則変換を行う（キャメルケース⇒スネークケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                String        文字列（キャメルケース）
' = 覚書    キャメルケース（ローワーキャメルケース）… getInputReader
'           スネークケース                          … get_input_reader
' ==================================================================
Public Function ConvCamelToSnake( _
    ByVal sInStr As String _
) As String
    Dim lLoopCnt As Long
    Dim sChar As String
    Dim sRetStr As String
    
    If sInStr = "" Then Exit Function
    
    sRetStr = ""
    For lLoopCnt = 1 To Len(sInStr)
        sChar = Mid$(sInStr, lLoopCnt, 1)
        If sChar = UCase(sChar) Then '大文字
            sRetStr = sRetStr & "_" & LCase(sChar)
        Else
            sRetStr = sRetStr & sChar
        End If
    Next lLoopCnt
    
    If Left(sRetStr, 1) = "_" Then
        sRetStr = Mid$(sRetStr, 2, Len(sRetStr))
    Else
        'Do Nothing
    End If
    
    ConvCamelToSnake = sRetStr
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
