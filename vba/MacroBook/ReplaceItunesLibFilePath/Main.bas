Attribute VB_Name = "Main"
Option Explicit

Public goPrgrsBar As New ProgressBar

Public Sub ReplaceItunesLibFilePath()
    
    Call BasicInfoInit
    Call ItunesInit
    
    Dim vOkCancel As Variant
    vOkCancel = MsgBox("�K���u�u�����v�A�u�u����v���Ƀt�@�C�������݂��Ă����ԂŖ{�c�[�������s���Ă��������I", vbOKCancel)
    If vOkCancel = vbCancel Then
        MsgBox "�v���O�������I�����܂�"
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
    MsgBox "�u�������I"
End Sub

'�f�o�b�O�p
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
    MsgBox "�o�͊����I"
End Sub


