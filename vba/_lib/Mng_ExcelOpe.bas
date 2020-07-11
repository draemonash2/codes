Attribute VB_Name = "Mng_ExcelOpe"
Option Explicit

' excel operation library v2.5

'************************************************************
'* �֐���`
'************************************************************
' ==================================================================
' = �T�v    �V�[�g�ꗗ�쐬
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Sub CreateSheetList()
    Dim oSheet As Object
    Dim lRowIdx As Long
    Dim lColumnIdx As Long
 
    If MsgBox("�A�N�e�B�u�Z�����牺�ɃV�[�g���ꗗ���쐬���Ă������ł����H", vbYesNo + vbDefaultButton2) = vbNo Then
        'None
    Else
        lRowIdx = ActiveCell.Row
        lColumnIdx = ActiveCell.Column
 
        For Each oSheet In ActiveWorkbook.Sheets
            Cells(lRowIdx, lColumnIdx).Value = oSheet.Name
            lRowIdx = lRowIdx + 1
        Next oSheet
    End If
End Sub

' ==================================================================
' = �T�v    ���[�N�V�[�g��V�K�쐬
' =         �d���������[�N�V�[�g������ꍇ�A_1, _2 ...�ƘA�ԂɂȂ�B
' =         �Ăяo�����ɂ͍쐬�������[�N�V�[�g����Ԃ��B
' = ����    sSheetName  [in]    String  �쐬����V�[�g��
' = �ߒl                        String  �쐬�����V�[�g��
' = �ˑ�    Mng_ExcelOpe.bas/ExistsWorksheet()
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Public Function CreateNewWorksheet( _
    ByVal sSheetName As String _
) As String
    Dim lShtIdx As Long
    
    lShtIdx = 0
    Dim bExistWorkSht As Boolean
    Do
        bExistWorkSht = ExistsWorksheet(sSheetName)
        If bExistWorkSht Then
            sSheetName = sSheetName & "_"
        Else
            lShtIdx = lShtIdx + 1 '�A�ԗp�̕ϐ�
        End If
    Loop While bExistWorkSht
    
    With ActiveWorkbook
        .Worksheets.Add(after:=.Worksheets(.Worksheets.Count)).Name = sSheetName
    End With
    CreateNewWorksheet = sSheetName
End Function

' ==================================================================
' = �T�v    �d������Worksheet���L�邩�`�F�b�N����B
' = ����    sTrgtShtName    [in]    String  �V�[�g��
' = �ߒl                            Boolean ���ݗL��
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Private Function ExistsWorksheet( _
    ByVal sTrgtShtName As String _
) As Boolean
    Dim lShtIdx As Long
    
    With ActiveWorkbook
        ExistsWorksheet = False
        For lShtIdx = 1 To .Worksheets.Count
            If .Worksheets(lShtIdx).Name = sTrgtShtName Then
                ExistsWorksheet = True
                Exit For
            End If
        Next
    End With
End Function

' ==================================================================
' = �T�v    �w�肵���L�[���[�h�̋߂��̃Z���l���擾����
' = ����    shTrgtSht       Worksheet   [in]    �ΏۃV�[�g
' = ����    sSearchKeyword  String      [in]    �����L�[���[�h
' = ����    lOffsetRow      Long        [in]    �s�I�t�Z�b�g
' = ����    lOffsetClm      Long        [in]    ��I�t�Z�b�g
' = ����    sOutputValue    String      [out]   �Z���l
' = �ߒl                    Boolean             �擾����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Public Function GetNearCellValue( _
    ByRef shTrgtSht As Worksheet, _
    ByVal sSearchKeyword As String, _
    ByVal lOffsetRow As Long, _
    ByVal lOffsetClm As Long, _
    ByRef sOutputValue As String _
) As Boolean
    With shTrgtSht
        Dim rFindResult As Range
        Set rFindResult = .Cells.Find(sSearchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            sOutputValue = ""
            GetNearCellValue = False
        Else
            If (rFindResult.Row + lOffsetRow) >= 1 And _
               (rFindResult.Column + lOffsetClm) >= 1 Then
                sOutputValue = .Cells( _
                                        rFindResult.Row + lOffsetRow, _
                                        rFindResult.Column + lOffsetClm _
                                    ).Value
                GetNearCellValue = True
            Else
                sOutputValue = ""
                GetNearCellValue = False
            End If
        End If
    End With
End Function
    Private Function Test_GetNearCellValue()
        Dim sSearchKeyword As String
        Dim lOffsetRow As Long
        Dim lOffsetClm As Long
        Dim sOutputValue As String
        Dim bRet As Boolean
        
        sSearchKeyword = "aaa"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 1, 1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, 1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        sSearchKeyword = "bbb"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -1, -1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -1, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -2, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -3, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -100, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, -100, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        sSearchKeyword = "ccc"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
    End Function

' ==================================================================
' = �T�v    �ΏۃV�[�g�̃Z������������B������Ȃ��ꍇ�A�����𒆒f����B
' = ����    shTrgtSht       Worksheet   [in]    �����ΏۃV�[�g
' = ����    sFindKeyword    String      [in]    �����ΏۃL�[���[�h
' = �ߒl                    Range               ��������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Public Function FindCell( _
    ByVal shTrgtSht As Worksheet, _
    ByVal sFindKeyword As String _
) As Range
    Set FindCell = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If FindCell Is Nothing Then
        MsgBox _
            "�Z����������Ȃ��������߁A�����𒆒f���܂��B" & vbNewLine & _
            "�@�����ΏۃV�[�g�F" & shTrgtSht.Name & vbNewLine & _
            "�@�����ΏۃL�[���[�h�F" & sFindKeyword, _
            vbCritical
        End
    End If
End Function
    Private Function Test_FindCell()
        Dim rFindResult As Range
        Debug.Print "*** test start!"
        Set rFindResult = FindCell(ActiveSheet, "�G�ۃ}�N��")
        Debug.Print "r" & rFindResult.Row & "c" & rFindResult.Column
        Set rFindResult = FindCell(ActiveSheet, "�G�ۃ}�N")
        Debug.Print "r" & rFindResult.Row & "c" & rFindResult.Column
        Debug.Print "*** test finish!"
    End Function



