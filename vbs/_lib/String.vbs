Option Explicit

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
' = 依存    なし
' = 所属    String.vbs
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim asSplitWord
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
    'Call Test_ExtractTailWord()
    Private Sub Test_ExtractTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
        Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
' = 依存    String.vbs/ExtractTailWord()
' = 所属    String.vbs
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim sTailWord
    Dim lRemoveLen
    
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
                RemoveTailWord = Left(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function
    'Call Test_RemoveTailWord()
    Private Sub Test_RemoveTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )    ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )        ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )          ' C:
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "", "\" )                 ' 
        Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )            ' a.txt（ファイル名かどうかは判断しない）
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )     ' C:\test\a.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        フォルダパス
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能
' = 依存    String.vbs/RemoveTailWord()
' = 所属    String.vbs
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "\" ) Then
        GetDirPath = RemoveTailWord( sFilePath, "\" )
    ElseIf InStr( sFilePath, "/" ) Then
        GetDirPath = RemoveTailWord( sFilePath, "/" )
    Else
        GetDirPath = sFilePath
    End If
End Function
    'Call Test_GetDirPath()
    Private Sub Test_GetDirPath()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetDirPath( "C:\test\a.txt" )    ' C:\test
        Result = Result & vbNewLine & GetDirPath( "http://test/a" )    ' http://test
        Result = Result & vbNewLine & GetDirPath( "C:_test_a.txt" )    ' C:_test_a.txt
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからファイル名を抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイル名
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能
' = 依存    String.vbs/ExtractTailWord()
' = 所属    String.vbs
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "\" ) Then
        GetFileName = ExtractTailWord( sFilePath, "\" )
    ElseIf InStr( sFilePath, "/" ) Then
        GetFileName = ExtractTailWord( sFilePath, "/" )
    Else
        GetFileName = sFilePath
    End If
End Function
    'Call Test_GetFileName()
    Private Sub Test_GetFileName()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileName( "C:\test\a.txt" )    ' a.txt
        Result = Result & vbNewLine & GetFileName( "http://test/a" )    ' a
        Result = Result & vbNewLine & GetFileName( "c:_test_a" )        ' c:_test_a
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスから拡張子を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        拡張子
' = 覚書    ・拡張子がない場合、空文字を返却する
' =         ・ファイル名も指定可能
' = 依存    String.vbs/ExtractTailWord()
' =         String.vbs/GetFileName()
' = 所属    String.vbs
' ==================================================================
Public Function GetFileExt( _
    ByVal sFilePath _
)
    Dim sFileName
    sFileName = GetFileName(sFilePath)
    If InStr(sFileName, ".") > 0 Then
        GetFileExt = ExtractTailWord(sFileName, ".")
    Else
        GetFileExt = ""
    End If
End Function
    'Call Test_GetFileExt()
    Private Sub Test_GetFileExt()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")     'txt
        Result = Result & vbNewLine & GetFileExt("c:\codes\test")         '
        Result = Result & vbNewLine & GetFileExt("test.txt")              'txt
        Result = Result & vbNewLine & GetFileExt("test")                  '
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
        Result = Result & vbNewLine & GetFileExt("test.aaa.txt")          'txt
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからファイルベース名を抽出する。
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイルベース名
' = 覚書    ・拡張子がない場合、空文字を返却する
' =         ・ファイル名も指定可能
' = 依存    String.vbs/RemoveTailWord()
' =         String.vbs/GetFileName()
' = 所属    String.vbs
' ==================================================================
Public Function GetFileBase( _
    ByVal sFilePath _
)
    Dim sFileName
    sFileName = GetFileName(sFilePath)
    GetFileBase = RemoveTailWord(sFileName, ".")
End Function
    'Call Test_GetFileBase()
    Private Sub Test_GetFileBase()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.txt")     'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test")         'test
        Result = Result & vbNewLine & GetFileBase("test.txt")              'test
        Result = Result & vbNewLine & GetFileBase("test")                  'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.aaa.txt") 'test.aaa
        Result = Result & vbNewLine & GetFileBase("test.aaa.txt")          'test.aaa
        MsgBox Result
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
' = 依存    String.vbs/GetDirPath()
' =         String.vbs/GetFileName()
' =         String.vbs/GetFileBase()
' =         String.vbs/GetFileExt()
' = 所属    String.vbs
' ==================================================================
Public Function GetFilePart( _
    ByVal sFilePath, _
    ByVal lPartType _
)
    Select Case lPartType
        Case 1: GetFilePart = GetDirPath(sFilePath)
        Case 2: GetFilePart = GetFileName(sFilePath)
        Case 3: GetFilePart = GetFileBase(sFilePath)
        Case 4: GetFilePart = GetFileExt(sFilePath)
        Case Else: GetFilePart = ""
    End Select
End Function
    'Call Test_GetFilePart()
    Private Sub Test_GetFilePart()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 0)     '
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 1)     'c:\codes
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 2)     'test.txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 3)     'test
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 4)     'txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 5)     '
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    日時形式を変換する。（例：2017/03/22 18:20:14 ⇒ 20170322-182014）
' = 引数    sDateTime   String  [in]  日時（YYYY/MM/DD HH:MM:SS）
' = 戻値                String        日時（YYYYMMDD-HHMMSS）
' = 覚書    主に日時をファイル名やフォルダ名に使用する際に使用する。
' = 依存    なし
' = 所属    String.vbs
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime _
)
    On Error Resume Next
    Dim sDateStr
    sDateStr = _
        String(4 - Len(Year(sDateTime)),   "0") & Year(sDateTime)   & _
        String(2 - Len(Month(sDateTime)),  "0") & Month(sDateTime)  & _
        String(2 - Len(Day(sDateTime)),    "0") & Day(sDateTime)    & _
        "-" & _
        String(2 - Len(Hour(sDateTime)),   "0") & Hour(sDateTime)   & _
        String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
        String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
    If Err.Number = 0 Then
        ConvDate2String = sDateStr
    Else
        ConvDate2String = ""
    End If
    On Error Goto 0
End Function
    'Call Test_ConvDate2String()
    Private Sub Test_ConvDate2String()
        MsgBox  ConvDate2String(Now()) & vbNewLine & _
                ConvDate2String("2001/12/32 1:00:0")
    End Sub

' ==================================================================
' = 概要    指定された文字列の文字列長（バイト数）を返却する
' = 引数    sInStr      String  [in]  文字列
' = 戻値                Long          文字列長（バイト数）
' = 覚書    標準で用意されている LenB() 関数は、Unicode における
' =         バイト数を返却するため半角文字も２文字としてカウントする。
' =           （例：LenB("ファイルサイズ ") ⇒ 16）
' =         そのため、半角文字を１文字としてカウントする本関数を用意。
' = 依存    なし
' = 所属    String.vbs
' ==================================================================
Public Function LenByte( _
    ByVal sInStr _
)
    Dim lIdx, sChar
    LenByte = 0
    If Trim(sInStr) <> "" Then
        For lIdx = 1 To Len(sInStr)
            sChar = Mid(sInStr, lIdx, 1)
            '２バイト文字は＋２
            If (Asc(sChar) And &HFF00) <> 0 Then
                LenByte = LenByte + 2
            Else
                LenByte = LenByte + 1
            End If
        Next
    End If
End Function
    'Call Test_LenByte()
    Private Sub Test_LenByte()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & LenByte( "aaa" )      ' 3
        Result = Result & vbNewLine & LenByte( "aaa " )     ' 4
        Result = Result & vbNewLine & LenByte( "" )         ' 0
        Result = Result & vbNewLine & LenByte( "あああ" )   ' 6
        Result = Result & vbNewLine & LenByte( "あああ " )  ' 7
        Result = Result & vbNewLine & LenByte( "ああ あ" )  ' 7
        Result = Result & vbNewLine & LenByte( Chr(9) )     ' 1
        Result = Result & vbNewLine & LenByte( Chr(10) )    ' 1
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    文字列長からタブ幅区切り位置までのタブ文字数を返却する
' = 引数    lLen            Long    [in]  文字列長
' = 引数    lLenMax         Long    [in]  文字列最大長
' = 引数    lTabWidth       Long    [in]  タブ文字幅
' = 引数    lPaddingTabNum  Long    [out] タブ文字数
' = 戻値                    Boolean       結果(OK,NG)
' = 覚書    実行例)lLen:3,lLenMax:9,lTabWidth:4
'             「xxx^^   ^   」
'             「xxxxxxxxx^  」
'               →return:3
' = 依存    String.vbs/CalcPaddingWidth()
' = 所属    String.vbs
' ==================================================================
Public Function CalcPaddingTabWidth( _
    ByVal lLen, _
    ByVal lLenMax, _
    ByVal lTabWidth, _
    ByRef lPaddingTabNum _
)
    Dim lPaddingLen
    Dim dCalcResult
    Dim lDigit
    If lTabWidth = 0 Then
        lPaddingTabNum = 0
        CalcPaddingTabWidth = False
    ElseIf lTabWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        lPaddingTabNum = 0
        CalcPaddingTabWidth = False
    Else
        'パディング幅(スペース)算出
        If CalcPaddingWidth(lLen, lLenMax, lTabWidth, lPaddingLen) Then
            'パディング幅(タブ)算出
            lDigit = 0
            dCalcResult = lPaddingLen / lTabWidth
            lPaddingTabNum = Fix((dCalcResult + (9 * (10 ^ (-1 * (lDigit + 1))))) * (10 ^ lDigit)) / (10 ^ lDigit)
            CalcPaddingTabWidth = True
        Else
            lPaddingTabNum = 0
            CalcPaddingTabWidth = False
        End If
    End If
End Function
    'Call Test_CalcPaddingTabWidth()
    Private Sub Test_CalcPaddingTabWidth()
        Dim Result
        Dim lPaddingLen
        Result = "[Result]"
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 1, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 4, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 6, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(4, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 2, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 3, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 4, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 5, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 19, 4, lPaddingLen)  & ":" & lPaddingLen '4
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 20, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 21, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 22, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 23, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 24, 4, lPaddingLen)  & ":" & lPaddingLen '6
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 15, 8, lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 16, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 17, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 18, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 19, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 20, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 21, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 22, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 23, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 24, 8, lPaddingLen)  & ":" & lPaddingLen '4
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, 5, 0,  lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, 5, -1, lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, -1, 4, lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(-1, 5, 4, lPaddingLen)  & ":" & lPaddingLen 'False
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    文字列長から区切り幅位置までの文字数を返却する
' = 引数    lLen        Long    [in]  文字列長
' = 引数    lLenMax     Long    [in]  文字列最大長
' = 引数    lSepWidth   Long    [in]  区切り幅
' = 引数    lPaddingLen Long    [out] 文字数
' = 戻値                Boolean       結果(OK,NG)
' = 覚書    実行例)lLen:3,lLenMax:9,lSepWidth:4
'             「xxx         」
'             「xxxxxxxxx   」
'               →return:9
' = 依存    なし
' = 所属    String.vbs
' ==================================================================
Public Function CalcPaddingWidth( _
    ByVal lLen, _
    ByVal lLenMax, _
    ByVal lSepWidth, _
    ByRef lPaddingLen _
)
    Dim lPaddingWidth
    If lSepWidth = 0 Then
        lPaddingLen = 0
        CalcPaddingWidth = False
    ElseIf lSepWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        lPaddingLen = 0
        CalcPaddingWidth = False
    Else
        If lLen > lLenMax Then
            lLenMax = lLen
        End If
        lPaddingWidth = lSepWidth - (lLenMax Mod lSepWidth)
        lPaddingLen = (lPaddingWidth + lLenMax) - lLen
        CalcPaddingWidth = True
    End If
End Function
    'Call Test_CalcPaddingWidth()
    Private Sub Test_CalcPaddingWidth()
        Dim Result
        Dim lPaddingLen
        Result = "[Result]"
        Result = Result & vbNewLine & CalcPaddingWidth(0, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & CalcPaddingWidth(3, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 5
        Result = Result & vbNewLine & CalcPaddingWidth(4, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(5, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(6, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(7, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(8, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(0, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & CalcPaddingWidth(1, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(2, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 6
        Result = Result & vbNewLine & CalcPaddingWidth(3, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 5
        Result = Result & vbNewLine & CalcPaddingWidth(4, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(5, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(6, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(7, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(8, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(0, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(3, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(4, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(5, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(6, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(7, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(8, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 6
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, 0, lPaddingLen)  & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, -1, lPaddingLen) & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(1, -1, 4, lPaddingLen) & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(-1, 5, 4, lPaddingLen) & ":" & lPaddingLen ' False
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    検索キーから値を抽出する。
' = 引数    sTrgtStr    String      [in]  対象文字列
' = 引数    asKeyList   String()    [in]  キーリスト
' = 引数    asValueList String()    [out] 値リスト
' = 引数    bTrimValue  Boolean     [in]  先頭末尾空白除去
' = 戻値                Boolean           一致有無
' = 覚書    ・asKeyListは必ず1オリジンで指定すること
' = 覚書    ・実行例は以下。
' =           ex1）
' =             <<入力>>
' =               sTrgtStr = "DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END: "
' =               asKeyList(1) = "FORM:"
' =               asKeyList(2) = "LSB:"
' =               asKeyList(3) = "DESC:"
' =               asKeyList(4) = "END:"
' =               bTrimValue = True
' =             <<出力>>
' =               asValueList(0) = ""
' =               asValueList(1) = "d AL"
' =               asValueList(2) = "12"
' =               asValueList(3) = "aBC"
' =               asValueList(4) = " "
' =           ex2）
' =             <<入力>>
' =               sTrgtStr = "DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END: "
' =               asKeyList(1) = "FORM:"
' =               asKeyList(2) = "ALLOC:"
' =               bTrimValue = False
' =             <<出力>>
' =               asValueList(0) = "aBC "
' =               asValueList(1) = "d AL LSB:12 "
' =               asValueList(2) = "aaa END: "
' = 依存    なし
' = 所属    String.vbs
' ==================================================================
Public Function ExtractValuesFrKeys( _
    ByVal sTrgtStr, _
    ByRef asKeyList(), _
    ByRef asValueList(), _
    ByVal bTrimValue _
)
    Dim lKeyNum
    lKeyNum = UBound(asKeyList)
    
    ReDim asValueList(lKeyNum) '0オリジン(0はヒットしなかった文字列)
    
    Dim alKeyMatchCharIdx()
    ReDim alKeyMatchCharIdx(lKeyNum)
    Dim lKeyIdx
    For lKeyIdx = 1 To lKeyNum
        alKeyMatchCharIdx(lKeyIdx) = 1
    Next
    
    Dim lCnfrmKeyIdxOld '0=未確定
    lCnfrmKeyIdxOld = 0
    
    Dim lCnfrmKeyStrPosOld
    lCnfrmKeyStrPosOld = 1
    
    ExtractValuesFrKeys = False
    
    Dim lCurStrPos
    For lCurStrPos = 1 To Len(sTrgtStr)
        Dim sCurChar
        sCurChar = Mid(sTrgtStr, lCurStrPos, 1)
        
        Dim lCnfrmKeyIdxNow '0=未確定
        lCnfrmKeyIdxNow = 0
        
        '*** 確定判定 ***
        For lKeyIdx = 1 To lKeyNum
            If asKeyList(lKeyIdx) = "" Then
                'Do Nothing
            Else
                Dim sCurKeyChar
                sCurKeyChar = Mid( _
                                    asKeyList(lKeyIdx), _
                                    alKeyMatchCharIdx(lKeyIdx), _
                                    1 _
                                )
                If sCurKeyChar = sCurChar Then
                    If Len(asKeyList(lKeyIdx)) <= alKeyMatchCharIdx(lKeyIdx) Then
                        lCnfrmKeyIdxNow = lKeyIdx
                        ExtractValuesFrKeys = True
                    Else
                        alKeyMatchCharIdx(lKeyIdx) = alKeyMatchCharIdx(lKeyIdx) + 1
                    End If
                Else
                    alKeyMatchCharIdx(lKeyIdx) = 1
                End If
                If lCnfrmKeyIdxNow > 0 Then
                    Exit For
                End If
            End If
        Next
        
        '*** 確定判定後事後処理 ***
        Dim lExtractLen
        If lCnfrmKeyIdxNow > 0 Then
            If lCnfrmKeyIdxOld > 0 Then 'ヒット二回目以降
                lExtractLen = lCurStrPos - Len(asKeyList(lCnfrmKeyIdxNow)) - lCnfrmKeyStrPosOld + 1
                asValueList(lCnfrmKeyIdxOld) = _
                    Mid( _
                        sTrgtStr, _
                        lCnfrmKeyStrPosOld, _
                        lExtractLen _
                    )
                If bTrimValue = True Then
                    asValueList(lCnfrmKeyIdxOld) = Trim(asValueList(lCnfrmKeyIdxOld))
                End If
            Else 'ヒット一回目
                If lCurStrPos > Len(asKeyList(lCnfrmKeyIdxNow)) Then
                    lExtractLen = lCurStrPos - Len(asKeyList(lCnfrmKeyIdxNow))
                    asValueList(0) = _
                        Mid( _
                            sTrgtStr, _
                            1, _
                            lExtractLen _
                        )
                    If bTrimValue = True Then
                        asValueList(0) = Trim(asValueList(0))
                    End If
                End If
            End If
            
            'クリア
            For lKeyIdx = 1 To lKeyNum
                alKeyMatchCharIdx(lKeyIdx) = 1
            Next
            
            '前回値更新
            lCnfrmKeyIdxOld = lCnfrmKeyIdxNow
            lCnfrmKeyStrPosOld = lCurStrPos + 1
        End If
    Next
    
    '最終要素取り出し
    If lCnfrmKeyIdxOld > 0 Then '1回以上ヒット
        lExtractLen = Len(sTrgtStr) - lCnfrmKeyStrPosOld + 1
        asValueList(lCnfrmKeyIdxOld) = _
            Mid( _
                sTrgtStr, _
                lCnfrmKeyStrPosOld, _
                lExtractLen _
            )
        If bTrimValue = True Then
            asValueList(lCnfrmKeyIdxOld) = Trim(asValueList(lCnfrmKeyIdxOld))
        End If
    Else '1回もヒットせず
        asValueList(0) = sTrgtStr
        If bTrimValue = True Then
            asValueList(0) = Trim(asValueList(0))
        End If
    End If
End Function
    Call Test_ExtractValuesFrKeys()
    Private Sub Test_ExtractValuesFrKeys()
        Dim asKeyList()
        Dim sOutMsg
        
        '正常系
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB;12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END: ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("d AL DESC:aBC ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("d AL DeSC:aBC ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB:12 DESC:DEf ALLOC:aaa END:", asKeyList, False, sOutMsg)
        
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, True, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        
        '異常系(asKeyList未初期化)
        'Dim asKeyList2()
        'Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList2, False, sOutMsg)
        
        '異常系(asKeyListに空白要素あり)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(3) = "LSB:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(3) = "LSB:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(3) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        
        '異常系(asKeyListが空白要素のみ)
        ReDim asKeyList(6)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        
        '異常系(sTrgtStr空文字列)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("", asKeyList, False, sOutMsg)
        
        '異常系(文字列長不足)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("a", asKeyList, False, sOutMsg)
        Msgbox sOutMsg
    End Sub
    Private Function TestSub_ExtractValuesFrKeysTrim( _
        ByVal sTrgtStr, _
        ByRef asKeyList(), _
        ByVal bTrimValue, _
        ByRef sOutMsg _
    )
        Dim asValueList()
        Dim bRet
        Redim asValueList(0)
        bRet = ExtractValuesFrKeys(sTrgtStr, asKeyList, asValueList, bTrimValue)
        sOutMsg = sOutMsg & vbNewLine & sTrgtStr
        sOutMsg = sOutMsg & vbNewLine & bRet
        
        Dim lIdx
        For lIdx = 0 To UBound(asValueList)
            If lIdx <= 0 Then
                sOutMsg = sOutMsg & vbNewLine & "other:""" & asValueList(lIdx) & """"
            Else
                sOutMsg = sOutMsg & vbNewLine & asKeyList(lIdx) & """" & asValueList(lIdx) & """"
            End If
        Next
        sOutMsg = sOutMsg & vbNewLine & ""
    End Function

'★TODO★実装
'Public Function ExtractRelativePath( _
'    ByVal sInFilePath As String, _
'    ByVal sMatchDirName As String, _
'    ByVal lRemeveDirLevel As Long, _
'    ByRef sRelativePath As String _
') As Boolean
