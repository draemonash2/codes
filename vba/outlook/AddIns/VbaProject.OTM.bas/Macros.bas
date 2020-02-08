Attribute VB_Name = "Macros"
Option Explicit

Public Sub 選択範囲のフラグをトグル()
    Dim objItem As Object
    Dim bFirstFlagCheck As Boolean
    Dim bUpdateFlagStatus
    bFirstFlagCheck = True
    For Each objItem In ActiveExplorer.Selection
        If bFirstFlagCheck = True Then
            If objItem.FlagStatus = olFlagMarked Then
                bUpdateFlagStatus = olNoFlag
            Else
                bUpdateFlagStatus = olFlagMarked
            End If
            bFirstFlagCheck = False
        Else
            'Do Nothing
        End If
        objItem.FlagStatus = bUpdateFlagStatus
        objItem.Save
    Next
End Sub

