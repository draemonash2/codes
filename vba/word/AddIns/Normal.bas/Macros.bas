Attribute VB_Name = "Macros"
Option Explicit

Public Sub 蛍光ペン黄()
    Call ToggleHighlighter(wdYellow)
End Sub

Public Sub 蛍光ペン緑()
    Call ToggleHighlighter(wdBrightGreen)
End Sub

'Wordのコマンドマクロ一覧を新規文書に出力する
Public Sub PrintOutWordCommandList()
    Dim wd As Word.Document
    Set wd = ThisDocument
    Application.ListCommands ListAllCommands:=True
    With ActiveDocument
        .PrintOut Background:=True, From:=1, To:=1 '1ページだけ出力されるように調整
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





