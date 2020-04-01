Attribute VB_Name = "Macros"
Option Explicit

Public Sub ŒuŒõƒyƒ“‰©()
    Call ToggleHighlighter(wdYellow)
End Sub

Public Sub ŒuŒõƒyƒ“—Î()
    Call ToggleHighlighter(wdBrightGreen)
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
