Option Explicit

' ==================================================================
' = 概要    デバッグ用の出力処理(Collection用)
' = 引数    cCollection     Collections [in]    出力対象
' = 戻値    なし
' = 覚書    ・出力後は処理を終了する
' = 依存    なし
' = 所属    Debug.vbs
' ==================================================================
Private Function DebugPrintClct( _
    ByRef cCollection _
)
    Dim sDebugMsg
    Dim vValue
    For Each vValue In cCollection
        sDebugMsg = sDebugMsg & vbNewLine & vValue
    Next
    MsgBox sDebugMsg
    WScript.Quit
End Function
