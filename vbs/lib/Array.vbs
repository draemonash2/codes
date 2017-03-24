Option Explicit

' ==================================================================
' = 概要    配列の中身をダイアログボックスに出力する。（デバッグ用）
' = 引数    asOutTrgtArray  String()    [in]    出力対象配列
' = 戻値    なし
' = 覚書    なし
' ==================================================================
Public Function OutputAllElement2Console( _
    ByRef asOutTrgtArray _
)
    Dim lIdx
    Dim sOutStr
    sOutStr = "EleNum :" & Ubound( asOutTrgtArray ) + 1
    For lIdx = 0 to UBound( asOutTrgtArray )
        sOutStr = sOutStr & vbNewLine & asOutTrgtArray(lIdx)
    Next
    WScript.Echo sOutStr
End Function

' ==================================================================
' = 概要    配列の中身をログファイルに出力する。（デバッグ用）
' = 引数    asOutTrgtArray  String()    [in]    出力対象配列
' = 戻値    なし
' = 覚書    ログファイル名は実行スクリプト名の拡張子を「.txt」に
' =         変えたものを出力する。
' ==================================================================
Public Function OutputAllElement2LogFile( _
    ByRef asOutTrgtArray _
)
    Dim lIdx
    Dim objLogFile
    Dim sLogFilePath
    Dim objWshShell
    
    sLogFilePath = Replace( WScript.ScriptFullName, ".vbs", ".log" )
    Set objLogFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sLogFilePath, 2, True )
    
    objLogFile.WriteLine "EleNum :" & Ubound( asOutTrgtArray ) + 1
    For lIdx = 0 to UBound( asOutTrgtArray )
        objLogFile.WriteLine asOutTrgtArray( lIdx )
    Next
    objLogFile.Close
    
    Set objWshShell = Nothing
    Set objLogFile = Nothing
End Function
'   Call Test_OutputAllElement2LogFile
    Private Sub Test_OutputAllElement2LogFile
        Dim asFileList()
        Redim asFileList(3)
        
        asFileList(0) = 1
        asFileList(1) = 0
        asFileList(2) = 1
        asFileList(3) = 0
    '   Call OutputAllElement2LogFile(asFileList)
        Call OutputAllElement2Console(asFileList)
    End Sub

' ==================================================================
' = 概要    定義済みの配列かどうかを判別する
' = 引数    asChkTrgtArray  String()    [in]    確認対象配列
' = 戻値                    Bool                結果（True:定義済み、False:未定義）
' = 覚書    配列でない場合、False が返却される。
' ==================================================================
Public Function IsArrayDefined( _
    ByRef asChkTrgtArray _
)
    Dim lArrayLastIdx
    On Error Resume Next
    lArrayLastIdx = UBound( asChkTrgtArray )
    If Err.Number <> 0 Then
        IsArrayDefined = False
        Err.Clear
    Else
        If lArrayLastIdx < 0 Then
            IsArrayDefined = False
        Else
            IsArrayDefined = True
        End If
    End If
    On Error Goto 0
End Function
'   Call Test_IsArrayDefined()
    Private Sub Test_IsArrayDefined()
        Dim Result
        Dim aTestArr01(0)
        Dim aTestArr02(1)
    '   Dim aTestArr03(-1) '定義できないのでテストしない
        Dim aTestArr04()
        ReDim aTestArr04(0)
        Dim aTestArr05()
        ReDim aTestArr05(1)
        Dim aTestArr06()
        ReDim aTestArr06(-1)
        Dim aTestArr07
        Set aTestArr07 = CreateObject("Scripting.FileSystemObject")
        Dim aTestArr08
        Dim aTestArr09()
        Result = "[Result]"
        Result = Result & vbNewLine & IsArrayDefined( aTestArr01 )  ' True
        Result = Result & vbNewLine & IsArrayDefined( aTestArr02 )  ' True
        Result = Result & vbNewLine & IsArrayDefined( aTestArr04 )  ' True
        Result = Result & vbNewLine & IsArrayDefined( aTestArr05 )  ' True
        Result = Result & vbNewLine & IsArrayDefined( aTestArr06 )  ' False
        Result = Result & vbNewLine & IsArrayDefined( aTestArr07 )  ' False
        Result = Result & vbNewLine & IsArrayDefined( aTestArr08 )  ' False
        Result = Result & vbNewLine & IsArrayDefined( aTestArr09 )  ' False
        MsgBox Result
    End Sub

