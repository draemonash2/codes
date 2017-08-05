Option Explicit

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
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
'   Call Test_ExtractTailWord()
    Private Sub Test_ExtractTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
        Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
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
'   Call Test_RemoveTailWord()
    Private Sub Test_RemoveTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )    ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )        ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )          ' C:
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "", "\" )                 ' 
        Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )            ' a.txt（ファイル名かどうかは判断しない）
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )     ' C:\test\a.txt
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        フォルダパス
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能
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
'   Call Test_GetDirPath()
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
'   Call Test_GetFileName()
    Private Sub Test_GetFileName()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileName( "C:\test\a.txt" )    ' a.txt
        Result = Result & vbNewLine & GetFileName( "http://test/a" )    ' a
        Result = Result & vbNewLine & GetFileName( "c:_test_a" )        ' c:_test_a
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスからファイル名（拡張子なし）を抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイル名（拡張子なし）
' = 覚書    拡張子が付与されていないファイルも存在する。そのため、
' =         "." が含まれていない場合も文字列を返却する。
' ==================================================================
Public Function GetFileBaseName( _
    ByVal sFilePath _
)
    GetFileBaseName = RemoveTailWord( ExtractTailWord( sFilePath, "\" ), "." )
End Function
'   Call Test_GetFileBaseName()
    Private Sub Test_GetFileBaseName()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileBaseName( "C:\test\a.txt" )    ' a
        Result = Result & vbNewLine & GetFileBaseName( "C:\test\a.t" )      ' a
        Result = Result & vbNewLine & GetFileBaseName( "C:\test\a." )       ' a
        Result = Result & vbNewLine & GetFileBaseName( "C:\test\a" )        ' a
        Result = Result & vbNewLine & GetFileBaseName( "C:\test\" )         ' 
        Result = Result & vbNewLine & GetFileBaseName( "C:\test" )          ' test
        Result = Result & vbNewLine & GetFileBaseName( "C:" )               ' C:
        Result = Result & vbNewLine & GetFileBaseName( "" )                 ' 
        Result = Result & vbNewLine & GetFileBaseName( "a.txt" )            ' a
        Result = Result & vbNewLine & GetFileBaseName( ".txt" )             ' 
        Result = Result & vbNewLine & GetFileBaseName( "a." )               ' a
        Result = Result & vbNewLine & GetFileBaseName( "." )                ' 
        Result = Result & vbNewLine & GetFileBaseName( "a" )                ' a
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定されたファイルパスから拡張子を抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        拡張子
' = 覚書    "." が含まれていない場合、空文字を返却する
' ==================================================================
Public Function GetFileExtName( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "." ) > 0 Then
        GetFileExtName = ExtractTailWord( sFilePath, "." )
    Else
        GetFileExtName = ""
    End If
End Function
'   Call Test_GetFileExtName()
    Private Sub Test_GetFileExtName()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileExtName( "C:\test\a.txt" ) ' txt
        Result = Result & vbNewLine & GetFileExtName( "C:\test\a.t" )   ' t
        Result = Result & vbNewLine & GetFileExtName( "C:\test\a." )    ' 
        Result = Result & vbNewLine & GetFileExtName( "C:\test\a" )     ' 
        Result = Result & vbNewLine & GetFileExtName( "C:\test\" )      ' 
        Result = Result & vbNewLine & GetFileExtName( "C:\test" )       ' 
        Result = Result & vbNewLine & GetFileExtName( "C:" )            ' 
        Result = Result & vbNewLine & GetFileExtName( "" )              ' 
        Result = Result & vbNewLine & GetFileExtName( "a.txt" )         ' txt
        Result = Result & vbNewLine & GetFileExtName( ".txt" )          ' txt
        Result = Result & vbNewLine & GetFileExtName( "a." )            ' 
        Result = Result & vbNewLine & GetFileExtName( "." )             ' 
        Result = Result & vbNewLine & GetFileExtName( "a" )             ' 
        MsgBox Result
    End Sub

' ==================================================================
' = 概要    指定された文字列の文字列長（バイト数）を返却する
' = 引数    sInStr      String  [in]  文字列
' = 戻値                Long          文字列長（バイト数）
' = 覚書    標準で用意されている LenB() 関数は、Unicode における
' =         バイト数を返却するため半角文字も２文字としてカウントする。
' =           （例：LenB("ファイルサイズ ") ⇒ 16）
' =         そのため、半角文字を１文字としてカウントする本関数を用意。
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
'   Call Test_LenByte()
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
' = 概要    日時形式を変換する。（例：2017/03/22 18:20:14 ⇒ 20170322-182014）
' = 引数    sDateTime   String  [in]  日時（YYYY/MM/DD HH:MM:SS）
' = 戻値                String        日時（YYYYMMDD-HHMMSS）
' = 覚書    主に日時をファイル名やフォルダ名に使用する際に使用する。
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime _
)
    ConvDate2String = Year(sDateTime) & _
                     String(2 - Len(Month(sDateTime)), "0") & Month(sDateTime) & _
                     String(2 - Len(Day(sDateTime)), "0") & Day(sDateTime) & _
                     "-" & _
                     String(2 - Len(Hour(sDateTime)), "0") & Hour(sDateTime) & _
                     String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
                     String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
End Function
    'Call Test_ConvDate2String()
    Private Sub Test_ConvDate2String()
        MsgBox ConvDate2String(Now())
    End Sub
