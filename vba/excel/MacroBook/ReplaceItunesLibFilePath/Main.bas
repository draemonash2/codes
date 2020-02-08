Attribute VB_Name = "Main"
Option Explicit

Public goPrgrsBar As New ProgressBar

Public Sub ReplaceItunesLibFilePath()
    
    Call BasicInfoInit
    Call ItunesInit
    
    Dim vOkCancel As Variant
    vOkCancel = MsgBox("必ず「置換元」、「置換先」共にファイルが存在している状態で本ツールを実行してください！", vbOKCancel)
    If vOkCancel = vbCancel Then
        MsgBox "プログラムを終了します"
        Exit Sub
    Else
        'Do Nothing
    End If
    
    Load goPrgrsBar
    goPrgrsBar.Show vbModeless
    Call BackUpItunesPlaylist
    Call GetBasicInfo
    Call ReplaceItunesLibLocation
    goPrgrsBar.Hide
    Unload goPrgrsBar
    Set goPrgrsBar = Nothing
    
    Call ItunesTerminate
    
    CreateObject("Wscript.Shell").Run """" & gsLogFilePath & """", 5
    MsgBox "置換完了！"
End Sub

'デバッグ用
Public Sub OutputItunesLibFilePath()
    
    Call ItunesInit
    
    Load goPrgrsBar
    goPrgrsBar.Show vbModeless
    
    Call OutputItunesLibLocation
    goPrgrsBar.Hide
    Unload goPrgrsBar
    Set goPrgrsBar = Nothing
    
    Call ItunesTerminate
    
    CreateObject("Wscript.Shell").Run """" & gsLogFilePath & """", 5
    MsgBox "出力完了！"
End Sub


