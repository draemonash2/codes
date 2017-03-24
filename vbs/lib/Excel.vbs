Option Explicit

'�}�N�����s
'�����͖��Ή�
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
    Call Test_ExecExcelMacro()
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
