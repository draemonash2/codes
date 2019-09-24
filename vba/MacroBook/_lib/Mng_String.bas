Attribute VB_Name = "Mng_String"
Option Explicit

' string manage library v1.5

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
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
    Private Sub Test_ExtractTailWord()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & ExtractTailWord("", "\")               '
        Result = Result & vbNewLine & ExtractTailWord("c:\a", "\")           ' a
        Result = Result & vbNewLine & ExtractTailWord("c:\a\", "\")          '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b", "\")         ' b
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\", "\")        '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\")   ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\\b\c.txt", "\")    ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\\b\c.txt", "\\") ' b\c.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
' = 依存    Mng_String.bas/ExtractTailWord()
' = 所属    Mng_String.bas
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
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & RemoveTailWord("", "\")               '
        Result = Result & vbNewLine & RemoveTailWord("c:\a", "\")           ' c:
        Result = Result & vbNewLine & RemoveTailWord("c:\a\", "\")          ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b", "\")         ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\", "\")        ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\")   ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\\b\c.txt", "\")    ' c:\\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\\b\c.txt", "\\") ' c:\a
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する
' = 引数    sFilePath       String  [in]  ファイルパス
' = 引数    bErrorEnable    Boolean [in]  エラー発生有効/無効(※)
' = 戻値                    Variant       フォルダパス
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能。
' =         (※) bErrorEnable にてファイルパス以外が指定された時の返却値を
' =         変えることが出来る｡
' =            True  : sFilePath を返却
' =            False : エラー値（xlErrNA）を返却
' = 依存    Mng_String.bas/RemoveTailWord()
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String, _
    Optional ByVal bErrorEnable As Boolean = False _
) As Variant
    If InStr(sFilePath, "\") Then
        GetDirPath = RemoveTailWord(sFilePath, "\")
    ElseIf InStr(sFilePath, "/") Then
        GetDirPath = RemoveTailWord(sFilePath, "/")
    Else
        If bErrorEnable = True Then
            GetDirPath = CVErr(xlErrNA)  'エラー値
        Else
            GetDirPath = sFilePath
        End If
    End If
End Function
    Private Sub Test_GetDirPath()
        Dim Result As String
        Dim vRet As Variant
        Result = "[Result]"
        vRet = GetDirPath("C:\test\a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' C:\test
        vRet = GetDirPath("http://test/a", True): Result = Result & vbNewLine & CStr(vRet)  ' http://test
        vRet = GetDirPath("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' エラー 2042
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:\test
        vRet = GetDirPath("http://test/a", False): Result = Result & vbNewLine & CStr(vRet) ' http://test
        vRet = GetDirPath("C:_test_a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:_test_a.txt
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:\test
        vRet = GetDirPath("http://test/a"): Result = Result & vbNewLine & CStr(vRet)        ' http://test
        vRet = GetDirPath("C:_test_a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:_test_a.txt
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからファイル名を抽出する
' = 引数    sFilePath       String  [in]  ファイルパス
' = 引数    bErrorEnable    Boolean [in]  エラー発生有効/無効(※)
' = 戻値                    Variant       ファイル名
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能。
' =         (※) bErrorEnable にてファイルパス以外が指定された時の返却値を
' =         変えることが出来る｡
' =            True  : sFilePath を返却
' =            False : エラー値（xlErrNA）を返却
' = 依存    Mng_String.bas/ExtractTailWord()
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String, _
    Optional ByVal bErrorEnable As Boolean = False _
) As Variant
    If InStr(sFilePath, "\") Then
        GetFileName = ExtractTailWord(sFilePath, "\")
    ElseIf InStr(sFilePath, "/") Then
        GetFileName = ExtractTailWord(sFilePath, "/")
    Else
        If bErrorEnable = True Then
            GetFileName = CVErr(xlErrNA)  'エラー値
        Else
            GetFileName = sFilePath
        End If
    End If
End Function
    Private Sub Test_GetFileName()
        Dim Result As String
        Dim vRet As Variant
        Result = "[Result]"
        vRet = GetFileName("C:\test\a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' a.txt
        vRet = GetFileName("http://test/a", True): Result = Result & vbNewLine & CStr(vRet)  ' a
        vRet = GetFileName("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' エラー 2042
        Result = Result & vbNewLine                                                          '
        vRet = GetFileName("C:\test\a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' a.txt
        vRet = GetFileName("http://test/a", False): Result = Result & vbNewLine & CStr(vRet) ' a
        vRet = GetFileName("C:_test_a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' c:_test_a
        Result = Result & vbNewLine                                                          '
        vRet = GetFileName("C:\test\a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' a.txt
        vRet = GetFileName("http://test/a"): Result = Result & vbNewLine & CStr(vRet)        ' a
        vRet = GetFileName("C:_test_a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' c:_test_a
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスから拡張子を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        拡張子
' = 覚書    ・拡張子がない場合、空文字を返却する
' =         ・ファイル名も指定可能
' = 依存    Mng_String.bas/ExtractTailWord()
' =         Mng_String.bas/GetFileName()
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetFileExt( _
    ByVal sFilePath As String _
) As String
    Dim sFileName As String
    sFileName = GetFileName(sFilePath)
    If InStr(sFileName, ".") > 0 Then
        GetFileExt = ExtractTailWord(sFileName, ".")
    Else
        GetFileExt = ""
    End If
End Function
    Private Sub Test_GetFileExt()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")     'txt
        Result = Result & vbNewLine & GetFileExt("c:\codes\test")         '
        Result = Result & vbNewLine & GetFileExt("test.txt")              'txt
        Result = Result & vbNewLine & GetFileExt("test")                  '
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
        Result = Result & vbNewLine & GetFileExt("test.aaa.txt")          'txt
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからファイルベース名を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイルベース名
' = 覚書    ・拡張子がない場合、空文字を返却する
' =         ・ファイル名も指定可能
' = 依存    Mng_String.bas/RemoveTailWord()
' =         Mng_String.bas/GetFileName()
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetFileBase( _
    ByVal sFilePath As String _
) As String
    Dim sFileName As String
    sFileName = GetFileName(sFilePath)
    GetFileBase = RemoveTailWord(sFileName, ".")
End Function
    Private Sub Test_GetFileBase()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.txt")     'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test")         'test
        Result = Result & vbNewLine & GetFileBase("test.txt")              'test
        Result = Result & vbNewLine & GetFileBase("test")                  'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.aaa.txt") 'test.aaa
        Result = Result & vbNewLine & GetFileBase("test.aaa.txt")          'test.aaa
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスから指定された一部を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 引数    lPartType   Long    [in]  抽出種別
' =                                     1) フォルダパス
' =                                     2) ファイル名
' =                                     3) ファイルベース名
' =                                     4) ファイル拡張子
' = 戻値                String        抽出した一部
' = 覚書    ・抽出種別が誤っている場合、空文字を返却する
' = 依存    Mng_String.bas/GetDirPath()
' =         Mng_String.bas/GetFileName()
' =         Mng_String.bas/GetFileBase()
' =         Mng_String.bas/GetFileExt()
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetFilePart( _
    ByVal sFilePath As String, _
    ByVal lPartType As Long _
) As String
    Select Case lPartType
        Case 1: GetFilePart = GetDirPath(sFilePath)
        Case 2: GetFilePart = GetFileName(sFilePath)
        Case 3: GetFilePart = GetFileBase(sFilePath)
        Case 4: GetFilePart = GetFileExt(sFilePath)
        Case Else: GetFilePart = ""
    End Select
End Function
    Private Sub Test_GetFilePart()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 0)     '
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 1)     'c:\codes
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 2)     'test.txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 3)     'test
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 4)     'txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 5)     '
        Debug.Print Result
    End Sub

' ==================================================================
' = 概要    日時形式を変換する。（例：2017/03/22 18:20:14 ⇒ 20170322-182014）
' = 引数    sDateTime   String  [in]  日時（YYYY/MM/DD HH:MM:SS）
' = 戻値                String        日時（YYYYMMDD-HHMMSS）
' = 覚書    主に日時をファイル名やフォルダ名に使用する際に使用する。
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime As String _
) As String
    ConvDate2String = Year(sDateTime) & _
                     String(2 - Len(Month(sDateTime)), "0") & Month(sDateTime) & _
                     String(2 - Len(Day(sDateTime)), "0") & Day(sDateTime) & _
                     "-" & _
                     String(2 - Len(Hour(sDateTime)), "0") & Hour(sDateTime) & _
                     String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
                     String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
End Function
    Private Sub Test_ConvDate2String()
        Debug.Print ConvDate2String(Now())
    End Sub


' ==================================================================
' = 概要    数字 型変換(String→Long)
' = 引数    sNum            String  [in]  数字(String型)
' = 戻値                    Long          数字(Long型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function
    Private Sub Test_NumConvStr2Lng()
        '★
    End Sub

' ==================================================================
' = 概要    数字 型変換(Long→String)
' = 引数    sNum            Long    [in]    数字(Long型)
' = 戻値                    String          数字(String型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function
    Private Sub Test_NumConvLng2Str()
        '★
    End Sub

' ==================================================================
' = 概要    正規表現検索を行う（Excel関数用）
' = 引数    sSearchPattern  String   [in]  検索パターン
' = 引数    sTargetStr      String   [in]  検索対象文字列
' = 引数    lMatchIdx       Long     [in]  検索結果インデックス（引数省略可）
' = 引数    bIsIgnoreCase   Boolean  [in]  大/小文字区別しないか（引数省略可）
' = 引数    bIsGlobal       Boolean  [in]  文字列全体を検索するか（引数省略可）
' = 戻値                    Variant        検索結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function RegExpSearch( _
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
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
    Private Sub Test_RegExpSearch()
        Dim sTargetStr As String
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print RegExpSearch(sTargetStr, " \w+\(")
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    正規表現検索を行う（Vbaマクロ関数用）
' = 引数    sTargetStr      String  [in]  検索対象文字列
' = 引数    sSearchPattern  String  [in]  検索パターン
' = 引数    oMatchResult    Object  [out] 検索結果
' = 戻値                    Boolean       ヒット有無
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ExecRegExp( _
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
    ByRef oMatchResult As Object, _
    Optional ByVal bIgnoreCase As Boolean = True, _
    Optional ByVal bGlobal As Boolean = True _
) As Boolean
    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.IgnoreCase = bIgnoreCase
    oRegExp.Global = bGlobal
    oRegExp.Pattern = sSearchPattern
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count = 0 Then
        ExecRegExp = False
    Else
        ExecRegExp = True
    End If
End Function
    Private Sub Test_ExecRegExp()
        Dim sTargetStr As String
        Dim oMatchResult As Object
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print ExecRegExp(sTargetStr, " \w+\(", oMatchResult)
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    命名規則変換を行う（スネークケース⇒パスカルケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                Variant       文字列（パスカルケース）
' = 覚書    ・スネークケースとパスカルケースについて
' =             - スネークケース … get_input_reader
' =             - パスカルケース … GetInputReader
' =               （＝アッパーキャメルケース）
' =         ・sInStr は単語のみを指定すること。（空白を含めない）
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ConvSnakeToPascal( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvSnakeToPascal = CVErr(xlErrName) 'エラー値
    Else
        sInStr = Replace(sInStr, "_", " ")
        sInStr = StrConv(sInStr, vbProperCase)
        sInStr = Replace(sInStr, " ", "")
        ConvSnakeToPascal = sInStr
    End If
End Function
    Private Sub Test_ConvSnakeToPascal()
        Debug.Print "*** test start! ***"
        Debug.Print ConvSnakeToPascal("get_input_reader") 'GetInputReader
        Debug.Print ConvSnakeToPascal("getinputreader")   'Getinputreader
        Debug.Print ConvSnakeToPascal("GetInputReader")   'GetInputReader
        Debug.Print ConvSnakeToPascal("get input reader") 'エラー 2029
        Debug.Print ConvSnakeToPascal("")                 '
        Debug.Print ConvSnakeToPascal("get_")             'Get
        Debug.Print ConvSnakeToPascal("_get_")            'Get
        Debug.Print ConvSnakeToPascal("get input_reader") 'エラー 2029
        Debug.Print ConvSnakeToPascal("get-input-reader") 'Get-input-reader
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    命名規則変換を行う（スネークケース⇒キャメルケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                Variant       文字列（キャメルケース）
' = 覚書    ・スネークケースとキャメルケースについて
' =             - スネークケース … get_input_reader
' =             - キャメルケース … getInputReader
' =               （＝ローワーキャメルケース）
' =         ・sInStr は単語のみを指定すること。（空白を含めない）
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ConvSnakeToCamel( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvSnakeToCamel = CVErr(xlErrName) 'エラー値
    Else
        sInStr = Replace(sInStr, "_", " ")
        sInStr = StrConv(sInStr, vbProperCase)
        sInStr = Replace(sInStr, " ", "")
        sInStr = LCase(Left$(sInStr, 1)) & _
                 Mid$(sInStr, 2, Len(sInStr))
        ConvSnakeToCamel = sInStr
    End If
End Function
    Private Sub Test_ConvSnakeToCamel()
        Debug.Print "*** test start! ***"
        Debug.Print ConvSnakeToCamel("get_input_reader") 'getInputReader
        Debug.Print ConvSnakeToCamel("getinputreader")   'getinputreader
        Debug.Print ConvSnakeToCamel("GetInputReader")   'getinputreader
        Debug.Print ConvSnakeToCamel("get input reader") 'エラー 2029
        Debug.Print ConvSnakeToCamel("")                 '
        Debug.Print ConvSnakeToCamel("get_")             'get
        Debug.Print ConvSnakeToCamel("_get_")            'get
        Debug.Print ConvSnakeToCamel("get input_reader") 'エラー 2029
        Debug.Print ConvSnakeToCamel("get-input-reader") 'get-input-reader
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    命名規則変換を行う（キャメルケース⇒スネークケース）
' = 引数    sInStr      String  [in]  文字列（スネークケース）
' = 戻値                Variant       文字列（キャメルケース）
' = 覚書    ・キャメルケースとスネークケースについて
' =             - キャメルケース … getInputReader
' =               （＝ローワーキャメルケース）
' =             - スネークケース … get_input_reader
' =         ・sInStr は単語のみを指定すること。（空白を含めない）
' =         ・アッパーキャメルケースも指定可能。
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ConvCamelToSnake( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvCamelToSnake = CVErr(xlErrName) 'エラー値
    Else
        If sInStr = "" Then
            ConvCamelToSnake = ""
        Else
            Dim lLoopCnt As Long
            Dim sChar As String
            Dim sRetStr As String
            
            sRetStr = ""
            For lLoopCnt = 1 To Len(sInStr)
                sChar = Mid$(sInStr, lLoopCnt, 1)
                If sChar <> "_" Then             '
                    If sChar = UCase(sChar) Then '大文字
                        If lLoopCnt = 1 Then     '一文字目
                            sRetStr = sRetStr & LCase(sChar)
                        Else
                            sRetStr = sRetStr & "_" & LCase(sChar)
                        End If
                    Else
                        sRetStr = sRetStr & sChar
                    End If
                Else
                    sRetStr = sRetStr & sChar
                End If
            Next lLoopCnt
            
            ConvCamelToSnake = sRetStr
        End If
    End If
End Function
    Private Sub Test_ConvCamelToSnake()
        Debug.Print "*** test start! ***"
        Debug.Print ConvCamelToSnake("getInputReader")  'get_input_reader
        Debug.Print ConvCamelToSnake("GetInputReader")  'get_input_reader
        Debug.Print ConvCamelToSnake("getInput_Reader") 'get_input__reader
        Debug.Print ConvCamelToSnake("GInput_Reader")   'g_input__reader
        Debug.Print ConvCamelToSnake("GINPUT")          'g_i_n_p_u_t
        Debug.Print ConvCamelToSnake("")                '
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    文字列を分割し、指定した要素の文字列を返却する
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 引数    iExtIndex   String  [in]  抽出する要素 ( 0 origin )
' = 戻値                Variant       抽出文字列
' = 覚書    iExtIndex が要素を超える場合、空文字列を返却する
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 概要    指定文字列の個数を返却する。
' = 引数    sTrgtStr      String  [in]  検索対象文字列
' = 引数    sSrchStr      String  [in]  検索文字列
' = 戻値                  Long          文字列の個数
' = 覚書    SplitStr との組み合わせでファイル名取り出しが可能。
' =           ex) B1 = C:\codes\c\Try04.c
' =               B2 = SplitStr( B1, "\", GetStrNum( B2, "\" ) )
' =                 ⇒ Try04.c
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function GetStrNum( _
    ByVal sTrgtStr As String, _
    ByVal sSrchStr As String _
) As Long
    Dim vSplitStr As Variant
    
    If sTrgtStr = "" Then
        GetStrNum = 0
    Else
        ' 文字列分割
        vSplitStr = Split(sTrgtStr, sSrchStr)
        GetStrNum = UBound(vSplitStr)
    End If
End Function
    Private Sub Test_GetStrNum()
        Debug.Print "*** test start! ***"
        Debug.Print GetStrNum("", "\")                '0
        Debug.Print GetStrNum("c:\a.txt", "\")        '1
        Debug.Print GetStrNum("c:\test\a.txt", "\")   '2
        Debug.Print GetStrNum("c:\test\a.txt", "")    '0
        Debug.Print GetStrNum("c:\test\\a.txt", "\")  '3
        Debug.Print GetStrNum("c:\test\\a.txt", "\\") '1
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    ビットＡＮＤ演算を行う。（数値）
' = 引数    cInVal1   Currency   [in]  入力値 左項（10進数数値）
' = 引数    cInVal2   Currency   [in]  入力値 右項（10進数数値）
' = 戻値              Variant          演算結果（10進数数値）
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 依存    Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/BitAndStrBin()
' =         Mng_String.bas/Bin2Hex()
' = 所属    Mng_String.bas
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
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 依存    Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/BitOrStrBin()
' =         Mng_String.bas/Bin2Hex()
' = 所属    Mng_String.bas
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
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 依存    Mng_String.bas/Dec2Hex()
' =         Mng_String.bas/BitShiftStrBin()
' =         Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/Bin2Hex()
' =         Mng_String.bas/Hex2Dec()
' = 所属    Mng_String.bas
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
' = 依存    Mng_String.bas/BitShiftStrBin()
' =         Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/Bin2Hex()
' = 所属    Mng_String.bas
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
' = 依存    Mng_String.bas/BitShiftLogStrBin()
' =         Mng_String.bas/BitShiftAriStrBin()
' = 所属    Mng_String.bas
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
' = 概要    基数変換 10進数→16進数(32bit用)
' = 引数    cInDecVal   Currency    [in]    10進数
' = 戻値                String              変換結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function Dec2Hex( _
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

' ==================================================================
' = 概要    基数変換 16進数→10進数(32bit用)
' = 引数    sInHexVal       String      [in]    16進数
' = 引数    bIsSignEnable   Boolean     [in]    符号有無
' = 戻値                    Variant             変換結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function Hex2Dec( _
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

' ==================================================================
' = 概要    基数変換 16進数→10進数
' = 引数    sHexVal         String      [in]    16進数
' = 戻値                    String              変換結果
' = 覚書    指定範囲以外の値を指定すると文字列 "error" を返却する。
' = 依存    Mng_String.bas/Hex2BinMap()
' = 所属    Mng_String.bas
' ==================================================================
Public Function Hex2Bin( _
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

' ==================================================================
' = 概要    基数変換 2進数→16進数
' = 引数    sBinVal         String      [in]    2進数
' = 引数    bIsUcase        Boolean     [in]    大文字小文字
' = 戻値                    String              変換結果
' = 覚書    指定範囲以外の値を指定すると文字列 "error" を返却する。
' = 依存    Mng_String.bas/Bin2HexMap()
' = 所属    Mng_String.bas
' ==================================================================
Public Function Bin2Hex( _
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

' ==================================================================
' = 概要    基数変換用マップ 16進数→2進数
' = 引数    sHexVal         String      [in]    16進数
' = 戻値                    String              変換結果
' = 覚書    指定範囲以外の値を指定すると文字列 "error" を返却する。
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function Hex2BinMap( _
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

' ==================================================================
' = 概要    基数変換用マップ 2進数→16進数
' = 引数    sBinVal         String      [in]    2進数
' = 引数    bIsUcase        Boolean     [in]    大文字小文字
' = 戻値                    String              変換結果
' = 覚書    指定範囲以外の値を指定すると文字列 "error" を返却する。
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function Bin2HexMap( _
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

' ==================================================================
' = 概要    論理ビットシフト（文字列版）
' = 引数    sInBinVal       String              [in]    2進数
' = 引数    lInShiftNum     Long                [in]    シフト数
' = 引数    eInDirection    E_SHIFT_DIRECTiON   [in]    シフト方向
' = 引数    lInBitLen       Long                [in]    ビットサイズ
' = 戻値                    String                      シフト結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function BitShiftLogStrBin( _
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

' ==================================================================
' = 概要    算術ビットシフト（文字列版）
' = 引数    sInBinVal           String              [in]    2進数
' = 引数    lInShiftNum         Long                [in]    シフト数
' = 引数    eInDirection        E_SHIFT_DIRECTiON   [in]    シフト方向
' = 引数    lInBitLen           Long                [in]    ビットサイズ
' = 引数    bIsSaveSignBit      Boolean             [in]    以下、参照
' = 引数    bIsExecAutoAlign    Boolean             [in]    以下、参照
' = 戻値                        String                      シフト結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function BitShiftAriStrBin( _
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

' ==================================================================
' = 概要    数字 型変換(String→Long)
' = 引数    sNum            String  [in]  数字(String型)
' = 戻値                    Long          数字(Long型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function

' ==================================================================
' = 概要    数字 型変換(Long→String)
' = 引数    lNum            Long    [in]    数字(Long型)
' = 戻値                    String          数字(String型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function

