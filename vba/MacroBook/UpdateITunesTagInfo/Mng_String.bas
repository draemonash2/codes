Attribute VB_Name = "Mng_String"
Option Explicit

' string manage library v1.1

' ==================================================================
' = 概要    フルパスから "ファイル名" を抽出する
' = 引数    sFullPath   String  [in]  フルパス
' = 戻値                Variant       ファイル名
' = 覚書    なし
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String _
) As Variant
    Dim asSplitWord() As String
    If InStr(sFilePath, "\") > 0 Then
        asSplitWord = Split(sFilePath, "\")
        GetFileName = asSplitWord(UBound(asSplitWord))
    Else
        GetFileName = CVErr(xlErrNA)  'エラー値
    End If
End Function

' ==================================================================x
' = 概要    フルパスから "ディレクトリパス" を抽出する
' = 引数    sFullPath   String  [in]  フルパス
' = 戻値                Variant       ディレクトリパス
' = 覚書    なし
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String _
) As Variant
    Dim asSplitWord() As String
    Dim sFileName As String
    If InStr(sFilePath, "\") > 0 Then
        asSplitWord = Split(sFilePath, "\")
        sFileName = asSplitWord(UBound(asSplitWord))
        GetDirPath = Replace( _
                                sFilePath, _
                                "\" & sFileName, _
                                "" _
                            )
    Else
        GetDirPath = CVErr(xlErrNA)  'エラー値
    End If
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

' ==================================================================
' = 概要    日時形式を変換する。（例：2017/03/22 18:20:14 ⇒ 20170322_182014）
' = 引数    sDateTime   String  [in]  日時（YYYY/MM/DD HH/MM/SS）
' = 戻値                String        日時（YYYYMMDD_HHMMSS）
' = 覚書    主に日時をファイル名やフォルダ名に使用する際に使用する。
' ==================================================================
Public Function JoinDateTime( _
    ByVal sDateTime As String _
) As String
    JoinDateTime = Year(sDateTime) & _
                     String(2 - Len(Month(sDateTime)), "0") & Month(sDateTime) & _
                     String(2 - Len(Day(sDateTime)), "0") & Day(sDateTime) & _
                     "_" & _
                     String(2 - Len(Hour(sDateTime)), "0") & Hour(sDateTime) & _
                     String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
                     String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
End Function
    Private Sub Test_JoinDateTime()
        MsgBox JoinDateTime(Now())
    End Sub
