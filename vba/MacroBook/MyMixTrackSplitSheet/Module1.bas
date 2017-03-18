Attribute VB_Name = "Module1"
Option Explicit

Const SHEET_NAME = "�����"
Enum ROW
    ROW_TITLE = 2
    ROW_START = 3
End Enum
Const ALBUM_NAME_COLUMN = 4
Private TARGET_SERVICE_COLUMN_TABLE() As Long

Public Function ExportInformationsInit()
    ReDim Preserve TARGET_SERVICE_COLUMN_TABLE(2)
    TARGET_SERVICE_COLUMN_TABLE(0) = 14  'Audacity
    TARGET_SERVICE_COLUMN_TABLE(1) = 18  'MixCloud
    TARGET_SERVICE_COLUMN_TABLE(2) = 20  'SuperTagEditer
End Function

Public Sub ExportInformations()
    Call ExportInformationsInit
    
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ThisWorkbook.Sheets(SHEET_NAME)
    
    '�t�H���_���w��
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFolderPath As String
    sOutputFolderPath = ShowFolderSelectDialog(objWshShell.SpecialFolders("Desktop"))

    '�t�@�C�����w��
    Dim sOutputFileBaseName As String
    sOutputFileBaseName = InputBox( _
                                "�t�@�C���x�[�X������͂��Ă�������", _
                                "test", _
                                shTrgtSht.Cells(ROW_START, ALBUM_NAME_COLUMN).Value _
                            )
    
    Dim lTrgtSrvcClmTblIdx As Long
    For lTrgtSrvcClmTblIdx = LBound(TARGET_SERVICE_COLUMN_TABLE) To UBound(TARGET_SERVICE_COLUMN_TABLE)
        Dim lClmIdx As Long
        lClmIdx = TARGET_SERVICE_COLUMN_TABLE(lTrgtSrvcClmTblIdx)
        
        '�ŏI�s����
        Dim lLastRow As Long
        Dim lRowIdx As Long
        lRowIdx = ROW_START
        Do While 1
            If shTrgtSht.Cells(lRowIdx, lClmIdx).Value = "" Then
                Exit Do
            Else
                'Do Nothing
            End If
            lRowIdx = lRowIdx + 1
        Loop
        lLastRow = lRowIdx - 1
        If lLastRow < ROW_START Then
            Exit For
        Else
            'Do Nothing
        End If
        
        '�t�@�C���p�X�쐬
        Dim sTitleName As String
        sTitleName = shTrgtSht.Cells(ROW_TITLE, lClmIdx).Value
        Dim sOutputFilePath As String
        sOutputFilePath = sOutputFolderPath & "\" & sOutputFileBaseName & "_" & sTitleName & ".txt"
        
        '�Z���͈͎擾
        Dim rTrgtRange As Range
        Set rTrgtRange = shTrgtSht.Range( _
                            shTrgtSht.Cells(ROW_START, lClmIdx), _
                            shTrgtSht.Cells(lLastRow, lClmIdx) _
                        )
        Dim asLine() As String
        Call ConvRange2Array(rTrgtRange, asLine(), True, Chr(9))
        
        '�Z���͈͏o��
        Open sOutputFilePath For Output As #1
        Dim lLineIdx As Long
        For lLineIdx = LBound(asLine) To UBound(asLine)
            Print #1, asLine(lLineIdx)
        Next lLineIdx
        Close #1
    Next lTrgtSrvcClmTblIdx
    
    '�{ Excel �t�@�C�����R�s�[
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile ThisWorkbook.FullName, sOutputFolderPath & "\" & sOutputFileBaseName & ".xlsm"
    
    MsgBox "�G�N�X�|�[�g�����I"
End Sub

' ==================================================================
' = �T�v    �Z���͈́iRange�^�j�𕶎���z��iString�z��^�j�ɕϊ�����B
' =         ��ɃZ���͈͂��e�L�X�g�t�@�C���ɏo�͂��鎞�Ɏg�p����B
' = ����    rCellsRange             Range   [in]  �Ώۂ̃Z���͈�
' = ����    asLine()                String  [out] ������ԊҌ�̃Z���͈�
' = ����    bIsInvisibleCellIgnore  String  [in]  ��\���Z���������s��
' = ����    sDelimiter              String  [in]  ��؂蕶��
' = �ߒl    �Ȃ�
' = �o��    �񂪗ׂ荇�����Z�����m�͎w�肳�ꂽ��؂蕶���ŋ�؂���
' ==================================================================
Private Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIsInvisibleCellIgnore As Boolean, _
    ByVal sDelimiter As String _
)
    Dim lLineIdx As Long
    lLineIdx = 0
    ReDim Preserve asLine(lLineIdx)
    
    Dim lRowIdx As Long
    For lRowIdx = 1 To rCellsRange.Rows.Count
        Dim lIgnoreCnt As Long
        lIgnoreCnt = 0
        Dim lClmIdx As Long
        For lClmIdx = 1 To rCellsRange.Columns.Count
            Dim sCurCellValue As String
            sCurCellValue = rCellsRange(lRowIdx, lClmIdx).Value
            '��\���Z���͖�������
            Dim bIsIgnoreCurExec As Boolean
            If bIsInvisibleCellIgnore = True Then
                If rCellsRange(lRowIdx, lClmIdx).EntireRow.Hidden = True Or _
                   rCellsRange(lRowIdx, lClmIdx).EntireColumn.Hidden = True Then
                    bIsIgnoreCurExec = True
                Else
                    bIsIgnoreCurExec = False
                End If
            Else
                bIsIgnoreCurExec = False
            End If
            
            If bIsIgnoreCurExec = True Then
                lIgnoreCnt = lIgnoreCnt + 1
            Else
                If lClmIdx = 1 Then
                    asLine(lLineIdx) = sCurCellValue
                Else
                    asLine(lLineIdx) = asLine(lLineIdx) & sDelimiter & sCurCellValue
                End If
            End If
        Next lClmIdx
        If lIgnoreCnt = rCellsRange.Columns.Count Then '��\���s�͍s���Z���Ȃ�
            'Do Nothing
        Else
            If lRowIdx = rCellsRange.Rows.Count Then '�ŏI�s�͍s���Z���Ȃ�
                'Do Nothing
            Else
                lLineIdx = lLineIdx + 1
                ReDim Preserve asLine(lLineIdx)
            End If
        End If
    Next lRowIdx
End Function

' ==================================================================
' = �T�v    �t�H���_�I���_�C�A���O��\������
' = ����    sInitPath   String  [in]  �f�t�H���g�t�H���_�p�X�i�ȗ��j
' = �ߒl                String        �t�H���_�I������
' = �o��    �Ȃ�
' ==================================================================
Public Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "�t�H���_��I�����Ă�������"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    '�_�C�A���O�\��
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ShowFolderSelectDialog = ""
    Else
        ShowFolderSelectDialog = fdDialog.SelectedItems.Item(1)
    End If
    
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        MsgBox ShowFolderSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") _
                )
    End Sub

