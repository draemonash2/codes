Attribute VB_Name = "IO_Button"
Option Explicit

Public Sub ChkButton_Click()
    Call ChkExec
End Sub

Public Sub WriTestRsltButton_Click()
    Call WriTestRsltExec
End Sub

Public Sub DocFilePathRefButton_Click()
    Call WriteDocFilePath
End Sub

Public Sub LogDirPathRefButton_Click()
    Call WriteLogDirPath
End Sub

Private Function WriteDocFilePath()
    Dim fdDialog As Office.FileDialog
    Dim asFilePath() As String
    Dim lResult As Long
    Dim lSelNum As Long
    Dim lSelIdx As Long
    
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイル選択ダイアログ"
    'fdDialog.InitialFileName = "D:\"
    fdDialog.AllowMultiSelect = False
    fdDialog.Filters.Add "全てのファイル", "*.*"
    
    'ダイアログ表示
    lResult = fdDialog.Show()
    
    If lResult <> -1 Then 'キャンセル押下
        ReDim Preserve asFilePath(0)
        asFilePath(0) = ""
    Else
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asFilePath(lSelNum - 1)
        For lSelIdx = 0 To lSelNum - 1
            asFilePath(lSelIdx) = fdDialog.SelectedItems(lSelIdx + 1)
        Next lSelIdx
        Call OutputDocFilePathCell(asFilePath(0))
    End If
    
    Set fdDialog = Nothing
End Function

Private Function WriteLogDirPath()
    Dim fdDialog As Office.FileDialog
    Dim sDirPath As String
    Dim lResult As Long
    
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダ選択ダイアログ"
    'fdDialog.InitialFileName = "D:\"
    'fdDialog.AllowMultiSelect = True 'フォルダは複数選択できない
    
    'ダイアログ表示
    lResult = fdDialog.Show()
    
    If lResult <> -1 Then 'キャンセル押下
        sDirPath = ""
    Else
        sDirPath = fdDialog.SelectedItems.Item(1)
        Call OutputLogDirPathCell(sDirPath)
    End If
    
    Set fdDialog = Nothing
End Function
