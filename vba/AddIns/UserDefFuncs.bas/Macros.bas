Attribute VB_Name = "Macros"
Option Explicit

' user define macros v2.2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' =============================================================================
' =  <<�}�N���ꗗ>>
' =    �E�I��͈͓��Œ���                   �I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' =
' =    �E�_�u���N�H�[�g�������ăZ���R�s�[   �_�u���N�I�[�e�[�V�����Ȃ��ŃZ���R�s�[����
' =    �E�I��͈͂��t�@�C���G�N�X�|�[�g     �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =    �E�I��͈͂��R�}���h���s             �I��͈͓��̃R�}���h�����s����B
' =
' =    �E�S�V�[�g�����R�s�[                 �u�b�N���̃V�[�g����S�ăR�s�[����
' =    �E�V�[�g�\����\����؂�ւ�         �V�[�g�\��/��\����؂�ւ���
' =    �E�V�[�g���בւ���Ɨp�V�[�g���쐬   �V�[�g���בւ���Ɨp�V�[�g�쐬
' =
' =    �E�Z�����̊ې������f�N�������g       �A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' =    �E�Z�����̊ې������C���N�������g     �@�`�M���w�肵�āA�w��ԍ��ȍ~���f�N�������g����
' =
' =    �E�c���[���O���[�v��                 �c���[�O���[�v������
' =    �E�n�C�p�[�����N�ꊇ�I�[�v��         �I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' =
' =    �E�t�H���g�F���g�O��                 �t�H���g�F���u�ԁv�́u�����v�Ńg�O������
' =    �E�w�i�F���g�O��                     �w�i�F���u���v�́u�w�i�F�Ȃ��v�Ńg�O������
' =
' =    �E�I�[�g�t�B�����s                   �I�[�g�t�B�������s����
' =============================================================================

'******************************************************************************
'* �萔��`
'******************************************************************************
'=== �ȉ��A�Z�����̊ې������f�N�������g()/�Z�����̊ې������C���N�������g() �p��` ===
Const NUM_MAX = 15
Const NUM_MIN = 1

'=== �ȉ��A�V�[�g���בւ���Ɨp�V�[�g���쐬() �p��` ===
Private Const WORK_SHEET_NAME = "�V�[�g���בւ���Ɨp"

Enum E_ROW
    ROW_BTN = 2
    ROW_TEXT_1 = 4
    ROW_TEXT_2
    ROW_SHT_NAME_TITLE = 7
    ROW_SHT_NAME_STRT
End Enum

Enum E_CLM
    CLM_BTN = 2
    CLM_SHT_NAME = 2
End Enum

Private Type T_SHORTCUT_KEY
    sKey As String
    sMacroName As String
End Type
Dim gtShortcutKey() As T_SHORTCUT_KEY

' *****************************************************************************
' * �V���[�g�J�b�g�L�[��`
' *****************************************************************************
Private Function InitUserDefShortcut()
    ' <<�V���[�g�J�b�g�L�[�ǉ����@>>
    '   (1) �ȉ��̒ǉ���ɁuAddUserDefShortcut()�v�Ăяo����ǉ�����B
    '       �������ɂ̓V���[�g�J�b�g�L�[�A�������Ƀ}�N�������w�肷��B
    '       �V���[�g�J�b�g�L�[�� Ctrl �� Shift �ȂǂƑg�ݍ��킹�Ďw��ł���B
    '         Shift�F+
    '         Ctrl �F^
    '         Alt  �F%
    '       �ڍׂ͈ȉ� URL �Q�ƁB
    '         https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '   (2) �}�N���u���[�U�[��`�V���[�g�J�b�g�L�[��ݒ�()�v�����s����B
    '
    ' <<�V���[�g�J�b�g�L�[�������@>>
    '   (1) �}�N���u���[�U�[��`�V���[�g�J�b�g�L�[������()�v�����s����B
    
    '������ �ǉ��� ������
'   Call AddUserDefShortcut("   ", "�I��͈͓��Œ���")

    Call AddUserDefShortcut("^+c", "�_�u���N�H�[�g�������ăZ���R�s�[")
'   Call AddUserDefShortcut("   ", "�I��͈͂��t�@�C���G�N�X�|�[�g")
'   Call AddUserDefShortcut("   ", "�I��͈͂��R�}���h���s")

'   Call AddUserDefShortcut("   ", "�S�V�[�g�����R�s�[")
'   Call AddUserDefShortcut("   ", "�V�[�g�\����\����؂�ւ�")
'   Call AddUserDefShortcut("   ", "�V�[�g���בւ���Ɨp�V�[�g���쐬")

'   Call AddUserDefShortcut("   ", "�Z�����̊ې������f�N�������g")
'   Call AddUserDefShortcut("   ", "�Z�����̊ې������C���N�������g")

'   Call AddUserDefShortcut("   ", "�c���[���O���[�v��")
'   Call AddUserDefShortcut("   ", "�n�C�p�[�����N�ꊇ�I�[�v��")

'   Call AddUserDefShortcut("   ", "�t�H���g�F���g�O��")
'   Call AddUserDefShortcut("   ", "�w�i�F���g�O��")
    
    Call AddUserDefShortcut("%^+{DOWN}", "'�I�[�g�t�B�����s(""Down"")'")
    Call AddUserDefShortcut("%^+{UP}", "'�I�[�g�t�B�����s(""Up"")'")
    Call AddUserDefShortcut("%^+{RIGHT}", "'�I�[�g�t�B�����s(""Right"")'")
    Call AddUserDefShortcut("%^+{LEFT}", "'�I�[�g�t�B�����s(""Left"")'")
    '������ �ǉ��� ������
End Function

Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[��ݒ�()
    Call InitUserDefShortcut
    Call EnableUserDefShortcut
End Sub

Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[������()
    Call InitUserDefShortcut
    Call DisableUserDefShortcut
End Sub

' *****************************************************************************
' * �}�N����`
' *****************************************************************************
' =============================================================================
' = �T�v�F�I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' =============================================================================
Public Sub �I��͈͓��Œ���()
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

' =============================================================================
' = �T�v�F�@�`�M���w�肵�āA�w��ԍ��ȍ~���f�N�������g����
' =============================================================================
Public Sub �Z�����̊ې������f�N�������g()
    Dim lTrgtNum As Long
    Dim sTrgtNum As String
    Dim lLoopCnt As Long
    
    sTrgtNum = InputBox("�f�N�������g���܂��B" & vbNewLine & "�J�n�ԍ�����͂��Ă��������B�i�A�`�N�j", "�ԍ�����", "")
    
    '���͒l�`�F�b�N
    If sTrgtNum = "" Then: MsgBox "���͒l�G���[�I": Exit Sub
    lTrgtNum = NumConvStr2Lng(sTrgtNum)
    If (lTrgtNum > NUM_MAX Or NUM_MIN + 1 > lTrgtNum) Then: MsgBox "���͒l�G���[�I": Exit Sub
    
    '�{����
    For lLoopCnt = lTrgtNum To NUM_MAX
        Selection.Replace _
            what:=NumConvLng2Str(lLoopCnt), _
            replacement:=NumConvLng2Str(lLoopCnt - 1)
    Next lLoopCnt
    MsgBox "�u�������I"
End Sub

' =============================================================================
' = �T�v�F�A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' =============================================================================
Public Sub �Z�����̊ې������C���N�������g()
    Dim lTrgtNum As Long
    Dim sTrgtNum As String
    Dim lLoopCnt As Long
    
    sTrgtNum = InputBox("�C���N�������g���܂��B" & vbNewLine & "�J�n�ԍ�����͂��Ă��������B�i�@�`�M�j", "�ԍ�����", "")
    
    '���͒l�`�F�b�N
    If sTrgtNum = "" Then: MsgBox "���͒l�G���[�I": Exit Sub
    lTrgtNum = NumConvStr2Lng(sTrgtNum)
    If (lTrgtNum > NUM_MAX - 1 Or NUM_MIN > lTrgtNum) Then: MsgBox "���͒l�G���[�I": Exit Sub
    
    '�{����
    For lLoopCnt = NUM_MAX - 1 To lTrgtNum Step -1
        Selection.Replace _
            what:=NumConvLng2Str(lLoopCnt), _
            replacement:=NumConvLng2Str(lLoopCnt + 1)
    Next lLoopCnt
    MsgBox "�u�������I"
End Sub

' =============================================================================
' = �T�v�F�u�b�N���̃V�[�g����S�ăR�s�[����
' = ���l�F�{�}�N�����G���[�ƂȂ�ꍇ�A�ȉ��̂����ꂩ�����{���邱�ƁB
' =       �E�c�[��->�Q�Ɛݒ� �ɂāuMicrosoft Forms 2.0 Object Library�v��I��
' =       �E�c�[��->�Q�Ɛݒ� ���́u�Q�Ɓv�ɂ� system32 ���́uFM20.DLL�v��I��
' =============================================================================
Public Sub �S�V�[�g�����R�s�[()
    Dim oSheet As Object
    Dim sSheetNames As String
    Dim doDataObj As New DataObject
    
    For Each oSheet In ActiveWorkbook.Sheets
        If sSheetNames = "" Then
            sSheetNames = oSheet.Name
        Else
            sSheetNames = sSheetNames + vbNewLine + oSheet.Name
        End If
    Next oSheet
    
    doDataObj.SetText sSheetNames
    doDataObj.PutInClipboard
    
    MsgBox "�u�b�N���̃V�[�g����S�ăR�s�[���܂���"
End Sub

' =============================================================================
' = �T�v�F�V�[�g�\��/��\����؂�ւ���
' =============================================================================
Public Sub �V�[�g�\����\����؂�ւ�()
    SheetVisibleSetting.Show
End Sub

' =============================================================================
' = �T�v�F�_�u���N�I�[�e�[�V�����Ȃ��ŃZ���R�s�[����
' =       ��\���Z���͖�������B�����͈͖͂��Ή��B
' =============================================================================
Public Sub �_�u���N�H�[�g�������ăZ���R�s�[()
    '*** ��\���Z���o�͔��� ***
    Dim bIsInvisibleCellIgnore As Boolean
    '���[�U�[�����P�������邽�߁A�f�t�H���g�Łu��\���Z�������v�Ƃ��Ă���
    bIsInvisibleCellIgnore = True
'    vAnswer = MsgBox("��\���Z���𖳎����܂����H", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "�����𒆒f���܂�"
'        End
'    End If
    
    '*** ��؂蕶������ ***
    Dim sDelimiter As String
    '���[�U�[�����P�������邽�߁A��Ԃ̋�؂蕶���̓f�t�H���g�Łu�^�u�����v�Œ�Ƃ��Ă���
    sDelimiter = Chr(9)
    
    '*** �Z���͈͂�String()�^�֕ϊ� ***
    Dim asLine() As String
    Call ConvRange2Array( _
        Selection, _
        asLine, _
        bIsInvisibleCellIgnore, _
        sDelimiter _
    )
    
    'String()�^�������N���b�v�{�[�h�ɃR�s�[
    Dim sBuf As String
    sBuf = ""
    Dim lLineIdx As Long
    For lLineIdx = LBound(asLine) To UBound(asLine)
        If lLineIdx = LBound(asLine) Then
            sBuf = asLine(lLineIdx)
        Else
            sBuf = sBuf & vbNewLine & asLine(lLineIdx)
        End If
    Next lLineIdx
    Call CopyText(sBuf)
    
    '�t�B�[�h�o�b�N
    Application.StatusBar = "���������������� �R�s�[�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v�F�I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =       �ׂ荇������̃Z���ɂ̓^�u������}�����ďo�͂���B
' =============================================================================
Public Sub �I��͈͂��t�@�C���G�N�X�|�[�g()
    Const TEMP_FILE_NAME As String = "ExportCellRange.tmp"
    
    '*** �Z���I�𔻒� ***
    If Selection.Count = 0 Then
        MsgBox "�Z�����I������Ă��܂���"
        MsgBox "�����𒆒f���܂�"
        End
    End If
    
    '*** Temp�t�@�C���Ǐo�� ***
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
        Dim sTmpPath As String
    sTmpPath = objWshShell.SpecialFolders("Templates") & "\" & TEMP_FILE_NAME
    Dim sDirPathOld As String
    Dim sFileNameOld As String
    If objFSO.FileExists(sTmpPath) Then
        Open sTmpPath For Input As #1
        Line Input #1, sDirPathOld
        Line Input #1, sFileNameOld
        Close #1
    Else
        sDirPathOld = objWshShell.SpecialFolders("Desktop")
        sFileNameOld = "export"
    End If
    
    '*** �t�H���_�p�X���� ***
    Dim sOutputDirPath As String
    sOutputDirPath = ShowFolderSelectDialog(sDirPathOld)
    If sOutputDirPath = "" Then
        MsgBox "�����ȃt�H���_���w��������̓t�H���_���I������܂���ł����B"
        MsgBox "�����𒆒f���܂��B"
        End
    Else
        'Do Nothing
    End If
    
    '*** �t�@�C�������� ***
    Dim sOutputFileName As String
    sOutputFileName = InputBox("�t�@�C��������͂��Ă��������B�i�g���q�Ȃ��j", "�t�@�C��������", sFileNameOld)
    
    '*** �t�@�C�����쐬 ***
    Dim sOutputFilePath As String
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName & ".txt"
    
    '*** �t�@�C���㏑������ ***
    If objFSO.FileExists(sOutputFilePath) Then
        Dim vAnswer As Variant
        vAnswer = MsgBox("�t�@�C�������݂��܂��B�㏑�����܂����H", vbOKCancel)
        If vAnswer = vbOK Then
            'Do Nothing
        Else
            MsgBox "�����𒆒f���܂��B"
            End
        End If
    Else
        'Do Nothing
    End If
    
    '*** ��\���Z���o�͔��� ***
    Dim bIsInvisibleCellIgnore As Boolean
    bIsInvisibleCellIgnore = True '���[�U�[�����P�������邽�߁A�f�t�H���g�Łu��\���Z�������v�Ƃ��Ă���
'    vAnswer = MsgBox("��\���Z���𖳎����܂����H", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "�����𒆒f���܂�"
'        End
'    End If
    
    '*** �t�@�C���o�͏��� ***
    Dim sDelimiter As String
    sDelimiter = Chr(9) '��Ԃ̋�؂蕶���́u�^�u�����v�Œ�
    
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                sDelimiter _
            )
    
    On Error Resume Next
    Open sOutputFilePath For Output As #1
    If Err.Number = 0 Then
        'Do Nothing
    Else
        MsgBox "�����ȃt�@�C���p�X���w�肳��܂���" & Err.Description
        MsgBox "�����𒆒f���܂��B"
        End
    End If
    On Error GoTo 0
    Dim lLineIdx As Long
    For lLineIdx = LBound(asRange) To UBound(asRange)
        Print #1, asRange(lLineIdx)
    Next lLineIdx
    Close #1
    
    '*** Temp�t�@�C�������o�� ***
    Open sTmpPath For Output As #1
    Print #1, sOutputDirPath
    Print #1, sOutputFileName
    Close #1
    
    MsgBox "�o�͊����I"
    
    '*** �o�̓t�@�C�����J�� ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v�F�I��͈͓��̃R�}���h�����s����B
' =       �P���I�����̂ݗL���B
' =============================================================================
Public Sub �I��͈͂��R�}���h���s()
    '*** �Z���I�𔻒� ***
    If Selection.Count = 0 Then
        MsgBox "�Z�����I������Ă��܂���"
        MsgBox "�����𒆒f���܂�"
        End
    End If
    
    '*** �͈̓`�F�b�N ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "�P���̂ݑI�����Ă�������"
        MsgBox "�����𒆒f���܂�"
        End
    End If
    
    '*** ��\���Z���o�͔��� ***
    Dim bIsInvisibleCellIgnore As Boolean
    bIsInvisibleCellIgnore = True '���[�U�[�����P�������邽�߁A�f�t�H���g�Łu��\���Z�������v�Ƃ��Ă���
'    vAnswer = MsgBox("��\���Z���𖳎����܂����H", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "�����𒆒f���܂�"
'        End
'    End If
    
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                "" _
            )
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\redirect.log"
    
    '*** �R�}���h���s ***
    Open sOutputFilePath For Append As #1
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Dim lLineIdx As Long
    For lLineIdx = LBound(asRange) To UBound(asRange)
        Print #1, asRange(lLineIdx)
        Print #1, ExecDosCmd(asRange(lLineIdx))
    Next lLineIdx
    Print #1, ""
    Close #1
    
    MsgBox "���s�����I"
    
    '*** �o�̓t�@�C�����J�� ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v�F�s���c���[�\���ɂ��ăO���[�v��
' = Usage�F�c���[�O���[�v���������͈͂�I�����A�}�N���u�c���[���O���[�v���v�����s����
' =============================================================================
Public Sub �c���[���O���[�v��()
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    '�O���[�v���ݒ�ύX
    ActiveSheet.Outline.SummaryRow = xlAbove
    
    lStrtRow = Selection(1).Row
    lLastRow = Selection(Selection.Count).Row
    lStrtClm = Selection(1).Column
    lLastClm = Selection(Selection.Count).Column
    
    '�O���[�v��
    Call TreeGroupSub( _
       ActiveSheet, _
       lStrtRow, _
       lLastRow, _
       lStrtClm, _
       lLastClm _
    )
End Sub

' =============================================================================
' = �T�v�F�V�[�g����ёւ���B
' =       �{���������s����ƁA�V�[�g���בւ���Ɨp�V�[�g���쐬����B
' =============================================================================
'���בւ��V�[�g ��Ɨp�V�[�g�쐬
Public Sub �V�[�g���בւ���Ɨp�V�[�g���쐬()
    Dim lShtIdx As Long
    Dim asShtName() As String
    Dim shWorkSht As Worksheet
    Dim bExistWorkSht As Boolean
    Dim lRowIdx As Long
    Dim lClmIdx As Long
    Dim lArrIdx As Long
    
    With ActiveWorkbook
        Application.ScreenUpdating = False

        ' === �V�[�g���擾 ===
        ReDim Preserve asShtName(.Worksheets.Count - 1)
        For lShtIdx = 1 To .Worksheets.Count
            asShtName(lShtIdx - 1) = .Sheets(lShtIdx).Name
        Next lShtIdx

        ' === ��Ɨp�V�[�g�쐬 ===
        bExistWorkSht = False
        For lShtIdx = 1 To .Worksheets.Count
            If .Sheets(lShtIdx).Name = WORK_SHEET_NAME Then
                bExistWorkSht = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lShtIdx
        If bExistWorkSht = True Then
            MsgBox "���Ɂu" & WORK_SHEET_NAME & "�v�V�[�g���쐬����Ă��܂��B"
            MsgBox "�����𑱂������ꍇ�́A�V�[�g���폜���Ă��������B"
            MsgBox "�����𒆒f���܂��B"
            End
        Else
            Set shWorkSht = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            shWorkSht.Name = WORK_SHEET_NAME
        End If

        '�V�[�g��񏑂�����
        shWorkSht.Cells(ROW_TEXT_1, CLM_SHT_NAME).Value = "��]�ʂ�ɃV�[�g������בւ��Ă��������B�i�ォ�珇�ɕ��בւ��܂��j"
        shWorkSht.Cells(ROW_TEXT_2, CLM_SHT_NAME).Value = "���בւ����I�������A�u���בւ����s�I�I�v�{�^���������Ă��������B"
        shWorkSht.Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).Value = "�V�[�g��"
        lArrIdx = 0
        For lRowIdx = ROW_SHT_NAME_STRT To ROW_SHT_NAME_STRT + UBound(asShtName)
            shWorkSht.Cells(lRowIdx, CLM_SHT_NAME).NumberFormatLocal = "@"
            shWorkSht.Cells(lRowIdx, CLM_SHT_NAME).Value = asShtName(lArrIdx)
            lArrIdx = lArrIdx + 1
        Next lRowIdx

        '�{�^���ǉ�
        With shWorkSht.Buttons.Add( _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Left, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Top, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Width, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Height _
        )
            .OnAction = "SortSheetPost"
            .Characters.Text = "���בւ����s�I�I"
        End With

        '�����ݒ�
        With ActiveSheet
            .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).Interior.ColorIndex = 34
            .Cells(ROW_BTN, CLM_BTN).RowHeight = 30
            .Cells(ROW_BTN, CLM_BTN).ColumnWidth = 40
            .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).HorizontalAlignment = xlCenter
            .Range( _
                .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME), _
                .Cells(.Rows.Count, CLM_SHT_NAME).End(xlUp) _
            ).Borders.LineStyle = True
            .Rows(ROW_SHT_NAME_TITLE + 1).Select
            ActiveWindow.FreezePanes = True
            .Rows(ROW_SHT_NAME_TITLE).Select
            Selection.AutoFilter
            .Cells(1, 1).Select
        End With
        
        Application.ScreenUpdating = True
    End With
End Sub

' =============================================================================
' = �T�v�F�V�[�g����ёւ���B
' =       �V�[�g���בւ���Ɨp�V�[�g�ɋL�ڂ̒ʂ�A�V�[�g����ёւ���B
' =       �K���V�[�g���בւ���Ɨp�V�[�g����Ăяo�����ƁI
' =============================================================================
Public Sub SortSheetPost()
    Dim asShtName() As String
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lArrIdx As Long
    Dim lRowIdx As Long
    
    With ActiveWorkbook
        '�V�[�g���擾
        lStrtRow = ROW_SHT_NAME_STRT
        lLastRow = .Sheets(WORK_SHEET_NAME).Cells(.Sheets(WORK_SHEET_NAME).Rows.Count, CLM_SHT_NAME).End(xlUp).Row
        ReDim Preserve asShtName(lLastRow - lStrtRow)
        lArrIdx = 0
        For lRowIdx = lStrtRow To lLastRow
            asShtName(lArrIdx) = .Sheets(WORK_SHEET_NAME).Cells(lRowIdx, CLM_SHT_NAME).Value
            lArrIdx = lArrIdx + 1
        Next lRowIdx
        
        '�V�[�g����r
        If UBound(asShtName) + 1 = .Sheets.Count - 1 Then
            'Do Nothing
        Else
            MsgBox "�V�[�g������v���܂���I"
            MsgBox "�����𒆒f���܂��B"
            End
        End If
        
        Application.ScreenUpdating = False
        
        '�V�[�g���בւ�
        For lArrIdx = 0 To UBound(asShtName)
            .Sheets(asShtName(lArrIdx)).Move Before:=Sheets(lArrIdx + 1)
        Next lArrIdx
        
        '��Ɨp�V�[�g�A�N�e�B�x�[�g
        .Sheets(WORK_SHEET_NAME).Activate
        
        '��Ɨp�V�[�g�폜�͎b�薳��
'        '��Ɨp�V�[�g�폜
'        Application.DisplayAlerts = False
'        .Sheets(WORK_SHEET_NAME).Delete
'        Application.DisplayAlerts = True
        
        Application.ScreenUpdating = True
    End With
    
    MsgBox "���בւ������I"
End Sub

' =============================================================================
' = �T�v�F�I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' =============================================================================
Public Sub �n�C�p�[�����N�ꊇ�I�[�v��()
    Dim Rng As Range
    
    If TypeName(Selection) = "Range" Then
        For Each Rng In Selection
            If Rng.Hyperlinks.Count > 0 Then Rng.Hyperlinks(1).Follow
        Next
    Else
        MsgBox "�Z���͈͂��I������Ă��܂���B", vbExclamation
    End If
End Sub

' =============================================================================
' = �T�v�F�t�H���g�F���u�ԁv�́u�����v�Ńg�O������
' =============================================================================
Public Sub �t�H���g�F���g�O��()
    Const COLOR_R As Long = 255
    Const COLOR_G As Long = 0
    Const COLOR_B As Long = 0
    If Selection(1).Font.Color = RGB(COLOR_R, COLOR_G, COLOR_B) Then
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Font.Color = RGB(COLOR_R, COLOR_G, COLOR_B)
    End If
End Sub

' =============================================================================
' = �T�v�F�w�i�F���u���v�́u�w�i�F�Ȃ��v�Ńg�O������
' =============================================================================
Public Sub �w�i�F���g�O��()
    Const COLOR_R As Long = 255
    Const COLOR_G As Long = 255
    Const COLOR_B As Long = 0
    If Selection(1).Interior.Color = RGB(COLOR_R, COLOR_G, COLOR_B) Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = RGB(COLOR_R, COLOR_G, COLOR_B)
    End If
End Sub

' =============================================================================
' = �T�v�F�I�[�g�t�B�������s����B
' =       �w�肵�������ɉ����đI��͈͂��L���ăI�[�g�t�B�������s����B
' =============================================================================
Public Sub �I�[�g�t�B�����s( _
    ByVal sDirection As String _
)
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    Dim lErrorNo As Long
    lErrorNo = 0
    
    Dim rSrc As Range
    Set rSrc = Selection
    Dim lSrcRow As Long
    Dim lSrcClm As Long
    lSrcRow = ActiveCell.Row
    lSrcClm = ActiveCell.Column
    
    '�I��͈͊g��
    If lErrorNo = 0 Then
        Select Case sDirection
            Case "Right": Range(Selection, Selection.Offset(0, 1)).Select
            Case "Left": Range(Selection, Selection.Offset(0, -1)).Select
            Case "Down": Range(Selection, Selection.Offset(1, 0)).Select
            Case "Up": Range(Selection, Selection.Offset(-1, 0)).Select
            Case Else: Debug.Assert 1
        End Select
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 1
        End If
    Else
        'Do Nothing
    End If
    
    '�I�[�g�t�B��
    If lErrorNo = 0 Then
        rSrc.AutoFill Destination:=Selection
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 2
        End If
    Else
        'Do Nothing
    End If
    
    '��ʃX�N���[��
    If lErrorNo = 0 Then
        Select Case sDirection
            Case "Right": Selection((lSrcRow - Selection(1).Row + 1), Selection.Columns.Count).Activate
            Case "Left": Selection((lSrcRow - Selection(1).Row + 1), 1).Activate
            Case "Down": Selection(Selection.Rows.Count, (lSrcClm - Selection(1).Column + 1)).Activate
            Case "Up": Selection(1, (lSrcClm - Selection(1).Column + 1)).Activate
'            Case "Right": Cells(lSrcRow, Selection.Columns.Count).Activate
'            Case "Left": Cells(lSrcRow, 1).Activate
'            Case "Down": Cells(Selection.Rows.Count, lSrcClm).Activate
'            Case "Up": Cells(1, lSrcClm).Activate
            Case Else: Debug.Assert 1
        End Select
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 3
        End If
    Else
        'Do Nothing
    End If
    
    Select Case lErrorNo
        Case 0: 'Do Nothing
        Case 1: Debug.Print "�y�I�[�g�t�B���W�J<" & sDirection & ">�z�ړ����G���[ No." & Err.Number & " : " & Err.Description
        Case 2: Debug.Print "�y�I�[�g�t�B���W�J<" & sDirection & ">�z�I�[�g�t�B�����G���[ No." & Err.Number & " : " & Err.Description
        Case 3: Debug.Print "�y�I�[�g�t�B���W�J<" & sDirection & ">�z�X�N���[�����G���[ No." & Err.Number & " : " & Err.Description
        Case Else: Debug.Assert 1
    End Select
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' *****************************************************************************
' * �����֐���`
' *****************************************************************************
Private Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function

Private Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function

Private Function TreeGroupSub( _
    ByRef shTrgtSht As Worksheet, _
    ByVal lGrpStrtRow As Long, _
    ByVal lGrpLastRow As Long, _
    ByVal lGrpStrtClm As Long, _
    ByVal lGrpLastClm As Long _
)
    Dim lCurRow As Long
    Dim lTrgtClm As Long
    Dim lAddRow As Long
    Dim lSubGrpStrtRow As Long
    Dim lSubGrpLastRow As Long
    Dim lSubGrpChkRow As Long
    
    Debug.Assert lGrpLastRow >= lGrpStrtRow
    Debug.Assert lGrpLastClm >= lGrpStrtClm
    
    If lGrpStrtClm >= lGrpLastClm Then
        'Do Nothing
    Else
        lCurRow = lGrpStrtRow
        lTrgtClm = lGrpStrtClm
        Do While lCurRow < lGrpLastRow
            If IsGroupParent(shTrgtSht, lCurRow, lTrgtClm) = True Then
                '=== �T�u�O���[�v�͈͔��� ===
                lSubGrpStrtRow = lCurRow + 1
                lSubGrpChkRow = lSubGrpStrtRow + 1
                Do While shTrgtSht.Cells(lSubGrpChkRow, lTrgtClm).Value = "" And _
                         lSubGrpChkRow <= lGrpLastRow
                    lSubGrpChkRow = lSubGrpChkRow + 1
                Loop
                lSubGrpLastRow = lSubGrpChkRow - 1
                '=== �T�u�O���[�v�̃O���[�v�� ===
                shTrgtSht.Range( _
                    shTrgtSht.Rows(lSubGrpStrtRow), _
                    shTrgtSht.Rows(lSubGrpLastRow) _
                ).Group
                '=== �ċA�Ăяo�� ===
                Call TreeGroupSub( _
                    shTrgtSht, _
                    lSubGrpStrtRow, _
                    lSubGrpLastRow, _
                    lTrgtClm + 1, _
                    lGrpLastClm _
                )
                lAddRow = lSubGrpLastRow - lSubGrpStrtRow + 1
            Else
                lAddRow = 1
            End If
            lCurRow = lCurRow + lAddRow
        Loop
    End If
End Function

' �w�肵���Z���̒����Z�����󔒂ŁA�E���Z�����󔒂łȂ��ꍇ�A
' �O���[�v�̐e�ł���Ɣ��f����B
Private Function IsGroupParent( _
    ByRef shTrgtSht As Worksheet, _
    ByVal lRow As Long, _
    ByVal lClm As Long _
) As Boolean
    Dim bRetVal As Boolean
    Dim sBtmCell As String
    Dim sBtmRightCell As String
    
    sBtmCell = ActiveSheet.Cells(lRow + 1, lClm + 0).Value
    sBtmRightCell = ActiveSheet.Cells(lRow + 1, lClm + 1).Value
    
    If sBtmCell = "" And sBtmRightCell <> "" Then     '�O���[�v�̐e
        bRetVal = True
    ElseIf sBtmCell <> "" And sBtmRightCell = "" Then '�O���[�v�̐e�łȂ�
        bRetVal = False
    Else                                              '����ȊO
        Debug.Assert 0 '���肦�Ȃ�
    End If
    
    IsGroupParent = bRetVal
End Function

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
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "�t�H���_��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j"
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
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists(sSelectedPath) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
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

'�R�}���h�����s
Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
    Private Sub Test_ExecDosCmd()
        Dim sBuf As String
        sBuf = sBuf & vbNewLine & ExecDosCmd("copy C:\Users\draem_000\Desktop\test.txt C:\Users\draem_000\Desktop\test2.txt")
        MsgBox sBuf
    End Sub

'�V���[�g�J�b�g�L�[��ǉ�
Private Function AddUserDefShortcut( _
    ByVal sKey As String, _
    ByVal sMacroName As String _
)
    If Sgn(gtShortcutKey) = 0 Then
        ReDim Preserve gtShortcutKey(0)
    Else
        ReDim Preserve gtShortcutKey(UBound(gtShortcutKey) + 1)
    End If
    gtShortcutKey(UBound(gtShortcutKey)).sKey = sKey
    gtShortcutKey(UBound(gtShortcutKey)).sMacroName = sMacroName
End Function

'�V���[�g�J�b�g�L�[��L����
Private Function EnableUserDefShortcut()
    Dim lIdx As Long
    For lIdx = LBound(gtShortcutKey) To UBound(gtShortcutKey)
        Application.OnKey gtShortcutKey(lIdx).sKey, gtShortcutKey(lIdx).sMacroName
    Next lIdx
End Function

'�V���[�g�J�b�g�L�[�𖳌���
Private Function DisableUserDefShortcut()
    Dim lIdx As Long
    For lIdx = LBound(gtShortcutKey) To UBound(gtShortcutKey)
        Application.OnKey gtShortcutKey(lIdx).sKey
    Next lIdx
End Function



