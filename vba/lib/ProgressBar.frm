VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrgrsBar 
   Caption         =   "UserForm1"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "PrgrsBar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "PrgrsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsgBlueBarMaxWidth As Single
Dim gbIsCanceled As Boolean
Dim glPercentOld As Long
 
Private Sub UserForm_Initialize()
    '### ユーザフォーム本体 ###
    With Me
        .Caption = "進捗状況"
    End With
    '### ラベル ###
    With LabelParcent
        .Caption = ""
        .Font.Bold = True
        .Font.Size = 12
    End With
    With LabelProcCmnt
        .Caption = ""
        .Font.Size = 10
    End With
    '### 青棒 ###
    With BlueBar
        .Width = 100
    End With
    gsgBlueBarMaxWidth = BlueBarFrame.Width - 2
    gbIsCanceled = False
    glPercentOld = 0
End Sub
 
Private Sub CancelButton_Click()
    gbIsCanceled = True
End Sub

Public Function Update( _
    ByVal lPercentCur As Long, _
    Optional ByVal sComment As String _
)
    Debug.Assert 0 <= lPercentCur And lPercentCur <= 100
    If glPercentOld = lPercentCur Then
        'Do Nothing
    Else
        LabelParcent.Caption = "処理中：" & lPercentCur & "%"
        LabelProcCmnt.Caption = sComment
        BlueBar.Width = gsgBlueBarMaxWidth * (lPercentCur / 100)
        DoEvents 'キャンセルボタンがない場合は「BackFrame.Repaint」を使用して処理を高速化すること
    End If
    glPercentOld = lPercentCur
End Function
 
Public Function IsCanceled() As Boolean
    IsCanceled = gbIsCanceled
End Function

