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
    fdDialog.Title = "�t�@�C���I���_�C�A���O"
    'fdDialog.InitialFileName = "D:\"
    fdDialog.AllowMultiSelect = False
    fdDialog.Filters.Add "�S�Ẵt�@�C��", "*.*"
    
    '�_�C�A���O�\��
    lResult = fdDialog.Show()
    
    If lResult <> -1 Then '�L�����Z������
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
    fdDialog.Title = "�t�H���_�I���_�C�A���O"
    'fdDialog.InitialFileName = "D:\"
    'fdDialog.AllowMultiSelect = True '�t�H���_�͕����I���ł��Ȃ�
    
    '�_�C�A���O�\��
    lResult = fdDialog.Show()
    
    If lResult <> -1 Then '�L�����Z������
        sDirPath = ""
    Else
        sDirPath = fdDialog.SelectedItems.Item(1)
        Call OutputLogDirPathCell(sDirPath)
    End If
    
    Set fdDialog = Nothing
End Function
