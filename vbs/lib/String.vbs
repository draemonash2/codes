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

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        フォルダパス
' = 覚書    なし
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath _
)
    GetDirPath = RemoveTailWord( sFilePath, "\" )
End Function

' ==================================================================
' = 概要    指定されたファイルパスからファイル名を抽出する
' = 引数    sFilePath   String  [in]  ファイルパス
' = 戻値                String        ファイル名
' = 覚書    なし
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath _
)
    GetFileName = ExtractTailWord( sFilePath, "\" )
End Function

' ==================================================================
' = 概要    指定された文字列の文字列長（バイト数）を返却する
' = 引数    sInStr      String  [in]  文字列
' = 戻値                Long          文字列長（バイト数）
' = 覚書    標準で用意されている LenB() 関数は、Unicode における
' =         バイト数を返却するため半角文字も２文字としてカウントする。
' =           （例：LenB("ファイルサイズ ") ⇒ 16）
' =         そのため、半角文字を１文字としてカウントする本関数を用意。
' ==================================================================
Function LenByte( _
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
