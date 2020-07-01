Attribute VB_Name = "Macros"
Option Explicit

Public Sub �u���y����()
    Call ToggleHighlighter(wdYellow)
End Sub

Public Sub �u���y����()
    Call ToggleHighlighter(wdBrightGreen)
End Sub

'Word�̃R�}���h�}�N���ꗗ��V�K�����ɏo�͂���
Public Sub PrintOutWordCommandList()
    Dim wd As Word.Document
    Set wd = ThisDocument
    Application.ListCommands ListAllCommands:=True
    With ActiveDocument
        .PrintOut Background:=True, From:=1, To:=1 '1�y�[�W�����o�͂����悤�ɒ���
        .Close SaveChanges:=wdDoNotSaveChanges
    End With
End Sub

Private Sub ToggleHighlighter( _
    ByVal wdTrgtColor As Variant _
)
    Dim wdColor As Variant
    If Selection.Range.HighlightColorIndex = wdTrgtColor Then
        wdColor = wdNoHighlight
    Else
        wdColor = wdTrgtColor
    End If
    Options.DefaultHighlightColorIndex = wdColor
    Selection.Range.HighlightColorIndex = wdColor
End Sub





