Option Explicit

'�}�N�����s
'�����͖��Ή�
' = �ˑ�    �Ȃ�
' = ����    Excel.vbs
Public Function ExecExcelMacro( _
    ByVal sExcelFilePath, _
    ByVal sMacroName _
)
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = False
    On Error Resume Next
    With objExcel
        .Workbooks.Open sExcelFilePath
        If Err.Number = 0 Then
            'Do Nothing
        Else
            WScript.Echo "�t�@�C����������܂���"
            WScript.Echo "�v���O�����𒆒f���܂�"
            WScript.Quit
        End If
        
        Dim sExcelFileName
        sExcelFileName = Mid( sExcelFilePath, InStrRev( sExcelFilePath, "\" ) + 1, len( sExcelFilePath ) )
        .Run "'" & sExcelFileName & "'!" & sMacroName
        If Err.Number = 0 Then
            'Do Nothing
        Else
            WScript.Echo "���s�ł��܂���ł���"
            WScript.Echo "�v���O�����𒆒f���܂�"
            WScript.Quit
        End If
        .Quit
    End With
    On Error Goto 0
    objExcel.Visible = True
    
    Set objExcel = Nothing
End Function
'   Call Test_ExecExcelMacro()
    Private Sub Test_ExecExcelMacro()
        Dim lTestCase
        lTestCase = InputBox("�e�X�g�P�[�X����͂��Ă�������")
        Select Case lTestCase
            Case 1:
                Call ExecExcelMacro( _
                    "C:\Users\draem_000\Desktop\test.xlsm", _
                    "testfunc" _
                )
            Case 2:
                Call ExecExcelMacro( _
                    "C:\Users\draem_000\Desktop\test2.xlsm", _
                    "testfunc" _
                )
            Case 3:
                Call ExecExcelMacro( _
                    "C:\Users\draem_000\Desktop\test.xlsm", _
                    "testfunc2" _
                )
            Case Else:
                MsgBox "�e�X�g�P�[�X������܂���"
        End Select
    End Sub

'�V�����G�N�Z���t�@�C�����쐬����
' = �ˑ�    �Ȃ�
' = ����    Excel.vbs
Public Function CreateNewExcelFile( _
    ByVal sBookPath _
)
    '[�Q�l] https://msdn.microsoft.com/ja-jp/library/office/ff198017.aspx
    Const xlExcel8 = 56
    Const xlWorkbookDefault = 51
    Const xlOpenXMLWorkbookMacroEnabled = 52
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists( sBookPath ) Then
        'Do Nothing
    Else
        Dim ExcelApp    ' �A�v���P�[�V����
        Dim ExcelBook   ' �u�b�N
        Set ExcelApp = CreateObject("Excel.Application")
        ExcelApp.DisplayAlerts = False
        
        ExcelApp.Workbooks.Add
        Set ExcelBook = ExcelApp.Workbooks( ExcelApp.Workbooks.Count )
        
        Dim sFileExt
        sFileExt = Mid( sBookPath, InStrRev( sBookPath, "." ) + 1, Len( sBookPath ) )
        on error resume next
        Select Case sFileExt
            Case "xlsx":    ExcelBook.SaveAs sBookPath, xlWorkbookDefault
            Case "xls":     ExcelBook.SaveAs sBookPath, xlExcel8
            Case "xlsm":    ExcelBook.SaveAs sBookPath, xlOpenXMLWorkbookMacroEnabled
            Case Else:      'Do Nothing
        End Select
        if Err.Number <> 0 then
            MsgBox( "ERROR:" & Err.Description )
        end if
        on error goto 0
        
        ExcelApp.Quit
        Set ExcelApp = Nothing
        ExcelApp = Empty
    End If
End Function
'   Call Test_CreateNewExcelFile()
    Private Sub Test_CreateNewExcelFile()
        Call CreateNewExcelFile("C:\Users\draem_000\Desktop\0.xlsx")
        Call CreateNewExcelFile("C:\Users\draem_000\Desktop\1.xlsm")
        Call CreateNewExcelFile("C:\Users\draem_000\Desktop\2.xls")
    End Sub

'Excel �V�[�g���
' = �ˑ�    �Ȃ�
' = ����    Excel.vbs
Public Function PrintExcelSheet( _
    ByVal sTrgtBookPath, _
    ByVal sTrgtSheetName, _
    ByVal lTrgtPageBegin, _
    ByVal lTrgtPageEnd, _
    ByVal lCopiesNum _
)
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    
    Dim objExcelBook
    Set objExcelBook = objExcel.Workbooks.Open( sTrgtBookPath )
    objExcel.Workbooks(1).Sheets( sTrgtSheetName ).Select
    objExcel.ActiveWindow.SelectedSheets.PrintOut lTrgtPageBegin, lTrgtPageEnd, lCopiesNum '���
    objExcel.Workbooks(1).Close
    Set objExcelBook = Nothing
    
    objExcel.quit()
    Set objExcel = Nothing
End Function
'   Call Test_PrintExcelSheet()
    Private Sub Test_PrintExcelSheet()
        Call PrintExcelSheet( _
            "C:\Users\draem_000\Desktop\�v�����^�C���N�l�܂�h�~�p����y�[�W.xlsx", _
            "Sheet2", _
            1, _
            1, _
            1 _
        )
    End Sub
