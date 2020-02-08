Attribute VB_Name = "Mng_Error"
Option Explicit
 
' error ribrary v1.01
 
'************************************************************
'* 構造体定義
'************************************************************
Public Enum E_ERROR_PROC
    ERROR_PROC_THROUGH = 0 'エラー出力後も無視して動作し続ける
    ERROR_PROC_STOP        'エラー出力後に停止する
End Enum
 
'************************************************************
'* モジュール内 変数定義
'************************************************************
Private gasErrorMsg() As String 'エラーメッセージ格納領域
 
'************************************************************
'* 関数定義
'************************************************************
' ==================================================================
' = 概要    初期化処理
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_Error.bas
' ==================================================================
Public Function ErrorMngInit()
    Erase gasErrorMsg
End Function
 
' ==================================================================
' = 概要    エラーメッセージ格納(出力しない)
' = 引数    sErrMsg     [in]    String  エラーメッセージ
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_Error.bas
' ==================================================================
Public Function StoreErrorMsg( _
    ByVal sErrMsg As String _
)
    Dim lErrMsgNum As Long
    If Sgn(gasErrorMsg) = 0 Then
        lErrMsgNum = 0
    Else
        lErrMsgNum = UBound(gasErrorMsg) + 1
    End If
    ReDim Preserve gasErrorMsg(lErrMsgNum)
    gasErrorMsg(lErrMsgNum) = sErrMsg
End Function
 
' ==================================================================
' = 概要    エラーメッセージ出力(格納したエラーを全て出力)
' = 引数    eErrorProc  [in]    E_ERROR_PROC    出力エラー処理
' = 戻値    なし
' = 依存    Mng_Error.bas/ExecuteErrorProc()
' = 所属    Mng_Error.bas
' ==================================================================
Public Function OutpErrorMsg( _
    ByVal eErrorProc As E_ERROR_PROC _
)
    Dim lErrMsgIdx As Long
    Dim sOutpMsg As String
 
    'エラー発生時のみ出力
    If Sgn(gasErrorMsg) = 0 Then
        'Do Nothing
    Else
        '#### エラー格納 ####
        sOutpMsg = ""
        For lErrMsgIdx = 0 To UBound(gasErrorMsg)
            sOutpMsg = sOutpMsg & _
                                "【ErrorNo." & lErrMsgIdx + 1 & "】" & vbCrLf & _
                                gasErrorMsg(lErrMsgIdx) & vbCrLf & vbCrLf
        Next lErrMsgIdx
 
        '#### エラー出力 ####
        If eErrorProc = ERROR_PROC_THROUGH Then
            MsgBox sOutpMsg, vbExclamation
        Else
            MsgBox sOutpMsg, vbCritical
        End If
        Call ErrorMngInit
 
        '#### エラー発生時処理 ####
        Call ExecuteErrorProc(eErrorProc)
    End If
 
End Function
 
' ==================================================================
' = 概要    エラー発生時に実行する処理を管理する。
' = 引数    eErrorProc  [in]    E_ERROR_PROC    出力エラー処理
' = 戻値    なし
' = 依存    ★/ChkExecTerminate()
' = 所属    Mng_Error.bas
' ==================================================================
Private Function ExecuteErrorProc( _
    ByVal eErrorProc As E_ERROR_PROC _
)
    Dim lProcSel As Long
    If eErrorProc = ERROR_PROC_THROUGH Then
        lProcSel = MsgBox("処理を継続しますか？", vbOKCancel)
        If lProcSel = vbOK Then
            MsgBox "処理を継続します！", vbExclamation
        Else
            MsgBox "処理を中断します！", vbCritical
            Call ChkExecTerminate
            End
        End If
    Else
        MsgBox "処理を中断します！", vbCritical
        Call ChkExecTerminate
        End
    End If
End Function

