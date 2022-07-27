Attribute VB_Name = "Mng_Excel_NewBook"
Option Explicit

Private gbIsNewExcelFIleOpened As Boolean
Private gbIsNewFIleOpenedJustBefore As Boolean

Public Function NewExcelMngInit()
    gbIsNewExcelFIleOpened = False
    gbIsNewFIleOpenedJustBefore = False
End Function

Public Function CreNewExcelFile() As Workbook
    Application.SheetsInNewWorkbook = 1
    Set CreNewExcelFile = Workbooks.Add
    gbIsNewFIleOpenedJustBefore = True
    gbIsNewExcelFIleOpened = True
End Function

Public Function SaveNewExcelFile( _
    ByRef wTrgtBook As Workbook, _
    ByRef sSaveTrgtFilePath As String _
)
    Debug.Assert gbIsNewExcelFIleOpened = True
    wTrgtBook.SaveAs Filename:=sSaveTrgtFilePath
    'wTrgtBook.Close
End Function

Public Function CopyRsltSht( _
    ByRef wTrgtBook As Workbook, _
    ByRef sShtName As String _
)
    ThisWorkbook.Sheets(sShtName).Visible = True
    ThisWorkbook.Sheets(sShtName).Copy After:=wTrgtBook.Worksheets(wTrgtBook.Worksheets.Count)
    ThisWorkbook.Sheets(sShtName).Visible = False
    
    'ブック作成時、先頭に不要シートが残っているので削除。
    If gbIsNewFIleOpenedJustBefore = True Then
        Application.DisplayAlerts = False
        wTrgtBook.Sheets(1).Delete
        Application.DisplayAlerts = True
        gbIsNewFIleOpenedJustBefore = False
    Else
        'Do Nothing
    End If
End Function

