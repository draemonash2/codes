Attribute VB_Name = "Macros"
Option Explicit

' user define macros v2.7

' =============================================================================
' =  <<�}�N���ꗗ>>
' =    �I��͈͓��Œ���                             �I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' =
' =    �_�u���N�H�[�g�������ăR�s�[                 �_�u���N�I�[�e�[�V�����Ȃ��ŃZ�����R�s�[����
' =    �_�u���N�H�[�g�������Ēǉ��R�s�[             �_�u���N�I�[�e�[�V�����Ȃ��ŃZ����ǉ��R�s�[����
' =
' =    �I��͈͂��t�@�C���G�N�X�|�[�g               �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =    �I��͈͂��܂Ƃ߂ăR�}���h���s               �I��͈͓��̃R�}���h���܂Ƃ߂Ď��s����B
' =    �I��͈͂����ꂼ��R�}���h���s               �I��͈͓��̃R�}���h�����ꂼ����s����B
' =    �I��͈͓��̌��������F��ύX                 �I��͈͓��̌��������F��ύX����
' =
' =    �S�V�[�g�����R�s�[                           �u�b�N���̃V�[�g����S�ăR�s�[����
' =    �V�[�g�\����\����؂�ւ�                   �V�[�g�\��/��\����؂�ւ���
' =    �V�[�g���בւ���Ɨp�V�[�g���쐬             �V�[�g���בւ���Ɨp�V�[�g�쐬
' =
' =    �Z�����̊ې������f�N�������g                 �A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' =    �Z�����̊ې������C���N�������g               �@�`�M���w�肵�āA�w��ԍ��ȍ~���f�N�������g����
' =
' =    �c���[���O���[�v��                           �c���[�O���[�v������
' =    �n�C�p�[�����N�ꊇ�I�[�v��                   �I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' =
' =    �t�H���g�F���g�O��                           �t�H���g�F���u�ԁv�́u�����v�Ńg�O������
' =    �w�i�F���g�O��                               �w�i�F���u���v�́u�w�i�F�Ȃ��v�Ńg�O������
' =
' =    �I�[�g�t�B�����s                             �I�[�g�t�B�������s����
' =    �A�N�e�B�u�Z���R�����g�̂ݕ\��               �A�N�e�B�u�Z���R�����g�̂ݕ\������
' =    �A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�     �A�N�e�B�u�Z���R�����g�̂ݕ\�����A�ړ�����
' =    �n�C�p�[�����N�Ŕ��                         �A�N�e�B�u�Z������n�C�p�[�����N��ɔ��
' =    MEMO�V�[�g�փW�����v                         �A�N�e�B�u�u�b�N��MEMO�V�[�g�ֈړ�����
' =
' =    Excel���ᎆ                                  Excel���ᎆ
' =    EpTree�̊֐��c���[��Excel�Ŏ�荞��          EpTree�̊֐��c���[��Excel�Ŏ�荞��
' =
' =    �����񕝒���                                 �񕝂�������������
' =    �����s������                                 �s����������������
' =
' =    �őO�ʂֈړ�                                 �őO�ʂֈړ�����
' =    �Ŕw�ʂֈړ�                                 �Ŕw�ʂֈړ�����
' =============================================================================

'******************************************************************************
'* �C���N���[�h
'******************************************************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'������Mng_Clipboard.bas/SetToClipboard()������
'Win32API�錾
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hData As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlag As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'�{���͂b����p�̕�����R�s�[�����A�Q�ڂ̈�����String�Ƃ��Ă���̂ŕϊ����s��ꂽ��ŃR�s�[�����B
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
'������Mng_Clipboard.bas/SetToClipboard()������

'******************************************************************************
'* �ݒ�l
'******************************************************************************
'=== �Z�����̊ې������f�N�������g()/�Z�����̊ې������C���N�������g() ===
Const NUM_MAX = 15
Const NUM_MIN = 1

'=== �A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�() ===
Const SETTING_KEY_CMNT_VSBL_ENB As String = "CMNT_VSBL_ENB"
Const SHTCUTKEY_KEYWORD_PREFIX As String = "SHTCUTKEY"

'=== �V�[�g���בւ���Ɨp�V�[�g���쐬() ===
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

' ==================================================================
' = �T�v    �V���[�g�J�b�g�L�[�ݒ���X�V����
' = ����    sOperate    String  [in] ������(Add/Update/Delete)
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/UpdateShtcutSetting()
' =         AddInSetting.cls/SearchWithKey()
' = ����    Macros.bas
' ==================================================================
Private Function UpdateShortcutKeySettings( _
    ByVal sOperate As String _
)
    ' <<�V���[�g�J�b�g�L�[�ǉ����@>>
    '   (1) �ȉ��̒ǉ���ɁuUpdateShtcutSetting()�v�Ăяo����ǉ�����B
    '       �������ɂ̓V���[�g�J�b�g�L�[�A�������Ƀ}�N�������w�肷��B
    '       �V���[�g�J�b�g�L�[�� Ctrl �� Shift �ȂǂƑg�ݍ��킹�Ďw��ł���B
    '         Shift�F+�ACtrl �F^�AAlt  �F%
    '       �ڍׂ͈ȉ� URL �Q�ƁB
    '         https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '   (2) �}�N���u���[�U�[��`�V���[�g�J�b�g�L�[��ݒ�()�v�����s����B
    '
    ' <<�V���[�g�J�b�g�L�[�������@>>
    '   (1) �}�N���u���[�U�[��`�V���[�g�J�b�g�L�[������()�v�����s����B
    
    '������ �ǉ��� ������
    Call UpdateShtcutSetting("", "�I��͈͓��Œ���", sOperate)
    
    Call UpdateShtcutSetting("^+c", "�_�u���N�H�[�g�������ăR�s�[", sOperate)
    Call UpdateShtcutSetting("^%c", "�_�u���N�H�[�g�������Ēǉ��R�s�[", sOperate)
    Call UpdateShtcutSetting("", "�I��͈͂��t�@�C���G�N�X�|�[�g", sOperate)
    Call UpdateShtcutSetting("", "�I��͈͂����ꂼ��R�}���h���s", sOperate)
    Call UpdateShtcutSetting("", "�I��͈͂��܂Ƃ߂ăR�}���h���s", sOperate)
    
    Call UpdateShtcutSetting("", "�S�V�[�g�����R�s�[", sOperate)
    Call UpdateShtcutSetting("", "�V�[�g�\����\����؂�ւ�", sOperate)
    Call UpdateShtcutSetting("", "�V�[�g���בւ���Ɨp�V�[�g���쐬", sOperate)
    
    Call UpdateShtcutSetting("", "�Z�����̊ې������f�N�������g", sOperate)
    Call UpdateShtcutSetting("", "�Z�����̊ې������C���N�������g", sOperate)
    
    Call UpdateShtcutSetting("", "�c���[���O���[�v��", sOperate)
    Call UpdateShtcutSetting("", "�n�C�p�[�����N�ꊇ�I�[�v��", sOperate)
    
    Call UpdateShtcutSetting("", "�t�H���g�F���g�O��", sOperate)
    Call UpdateShtcutSetting("", "�w�i�F���g�O��", sOperate)
    
    Call UpdateShtcutSetting("", "Excel���ᎆ", sOperate)
    Call UpdateShtcutSetting("", "EpTree�̊֐��c���[��Excel�Ŏ�荞��", sOperate)
    
    Call UpdateShtcutSetting("%^+{DOWN}", "'�I�[�g�t�B�����s(""Down"")'", sOperate)
    Call UpdateShtcutSetting("%^+{UP}", "'�I�[�g�t�B�����s(""Up"")'", sOperate)
    Call UpdateShtcutSetting("%^+{RIGHT}", "'�I�[�g�t�B�����s(""Right"")'", sOperate)
    Call UpdateShtcutSetting("%^+{LEFT}", "'�I�[�g�t�B�����s(""Left"")'", sOperate)
    
    Call UpdateShtcutSetting("%{F9}", "�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�_���[�h�ؑ�", sOperate)
    
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim sValue As String
    Dim bIsRet As Boolean
    bIsRet = clSetting.SearchWithKey(SETTING_KEY_CMNT_VSBL_ENB, sValue)
    If bIsRet = True Then
        If sValue = "True" Then
            Call UpdateShtcutSetting("{DOWN}", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Down"")'", sOperate)
            Call UpdateShtcutSetting("{UP}", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Up"")'", sOperate)
            Call UpdateShtcutSetting("{RIGHT}", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Right"")'", sOperate)
            Call UpdateShtcutSetting("{LEFT}", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Left"")'", sOperate)
        ElseIf sValue = "False" Then
            Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Down"")'", sOperate)
            Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Up"")'", sOperate)
            Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Right"")'", sOperate)
            Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Left"")'", sOperate)
        Else
            Debug.Assert False
        End If
    Else
        Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Down"")'", sOperate)
        Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Up"")'", sOperate)
        Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Right"")'", sOperate)
        Call UpdateShtcutSetting("", "'�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�(""Left"")'", sOperate)
    End If
    
    Call UpdateShtcutSetting("^+j", "�n�C�p�[�����N�Ŕ��", sOperate)
    Call UpdateShtcutSetting("^%{HOME}", "MEMO�V�[�g�փW�����v", sOperate)
    
    Call UpdateShtcutSetting("", "Excel���ᎆ", sOperate)
    
    Call UpdateShtcutSetting("", "�����񕝒���", sOperate)
    Call UpdateShtcutSetting("", "�����s������", sOperate)
    
    Call UpdateShtcutSetting("^+f", "�őO�ʂֈړ�", sOperate)
    Call UpdateShtcutSetting("^+b", "�Ŕw�ʂֈړ�", sOperate)
    '������ �ǉ��� ������
End Function

' *****************************************************************************
' * �O�����J�p�}�N��
' *****************************************************************************
' =============================================================================
' = �T�v    �I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͓��Œ���()
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

' =============================================================================
' = �T�v    �@�`�M���w�肵�āA�w��ԍ��ȍ~���f�N�������g����
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = ����    Macros.bas
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
' = �T�v    �A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = ����    Macros.bas
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
' = �T�v    �u�b�N���̃V�[�g����S�ăR�s�[����
' = �o��    �{�}�N�����G���[�ƂȂ�ꍇ�A�ȉ��̂����ꂩ�����{���邱�ƁB
' =           �E�c�[��->�Q�Ɛݒ� �ɂāuMicrosoft Forms 2.0 Object Library�v��I��
' =           �E�c�[��->�Q�Ɛݒ� ���́u�Q�Ɓv�ɂ� system32 ���́uFM20.DLL�v��I��
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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
' = �T�v    �V�[�g�\��/��\����؂�ւ���
' = �o��    �Ȃ�
' = �ˑ�    SheetVisibleSetting.cls/SheetVisibleSetting()
' = ����    Macros.bas
' =============================================================================
Public Sub �V�[�g�\����\����؂�ւ�()
    SheetVisibleSetting.Show
End Sub

' =============================================================================
' = �T�v    �_�u���N�I�[�e�[�V�����Ȃ��ŃZ���R�s�[����
' =         ��\���Z���͖�������B�����͈͖͂��Ή��B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' = ����    Macros.bas
' =============================================================================
Public Sub �_�u���N�H�[�g�������ăR�s�[()
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
    Call SetToClipboard(sBuf)
    
    '�t�B�[�h�o�b�N
    Application.StatusBar = "���������������� �R�s�[�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' ==================================================================
' = �T�v    �I��͈͂��N���b�v�{�[�h�֒ǉ��R�s�[
' =         �_�u���N�I�[�e�[�V�����Ȃ��ŃZ���R�s�[����
' =         ��\���Z���͖�������B�����͈͖͂��Ή��B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' = ����    Macros.bas
' ==================================================================
Public Sub �_�u���N�H�[�g�������Ēǉ��R�s�[()
    Dim bIsInvisibleCellIgnore As Boolean
    Dim sDelimiter As String
    bIsInvisibleCellIgnore = True
    sDelimiter = Chr(9)
    
    '*** �����e�L�X�g�擾 ***
    Dim sOrgText As String
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        sOrgText = .GetText
    End With
    
    '*** �ǉ��e�L�X�g�擾 ***
    Dim asLine() As String
    Call ConvRange2Array( _
        Selection, _
        asLine, _
        bIsInvisibleCellIgnore, _
        sDelimiter _
    )
    
    Dim sNewText As String
    sNewText = ""
    Dim lLineIdx As Long
    For lLineIdx = LBound(asLine) To UBound(asLine)
        If lLineIdx = LBound(asLine) Then
            sNewText = asLine(lLineIdx)
        Else
            sNewText = sNewText & vbNewLine & asLine(lLineIdx)
        End If
    Next lLineIdx
    
    '*** �N���b�v�{�[�h�ݒ� ***
    Dim sOutText As String
    sOutText = sOrgText & vbNewLine & sNewText
    Call SetToClipboard(sOutText)

    '*** �t�B�[�h�o�b�N ***
    Application.StatusBar = "���������������� �ǉ��R�s�[�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v    �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =         �ׂ荇������̃Z���ɂ̓^�u������}�����ďo�͂���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_Array.bas/ConvRange2Array()
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͂��t�@�C���G�N�X�|�[�g()
    Const TEMP_FILE_NAME As String = "ExportCellRange.tmp"
    Const FILE_EXTENTION As String = "csv"
    Const DELIMITER As String = ","
    
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
    sTmpPath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Temp\" & TEMP_FILE_NAME
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
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName & "." & FILE_EXTENTION
    
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
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                DELIMITER _
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
' = �T�v    �I��͈͓��̃R�}���h���o�b�`�t�@�C���ɏ����o���Ă܂Ƃ߂Ď��s����B
' =         �P���I�����̂ݗL���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_SysCmd.bas/ExecDosCmd()
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͂��܂Ƃ߂ăR�}���h���s()
    Const BAT_FILE_NAME As String = "command.bat"
    
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
    
    Dim sBatFileDirPath As String
    Dim sBatFilePath As String
    sBatFileDirPath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Temp"
    sBatFilePath = sBatFileDirPath & "\" & BAT_FILE_NAME
    
    Call OutputTxtFile(sBatFilePath, asRange)
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\redirect.log"
    
    '*** �R�}���h���s ***
    Open sOutputFilePath For Append As #1
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Print #1, ExecDosCmd(sBatFilePath)
    Print #1, ""
    Close #1
    
    '*** �o�b�`�t�@�C���폜 ***
    Kill sBatFilePath
    
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
' = �T�v    �I��͈͓��̃R�}���h�����ꂼ����s����B
' =         �P���I�����̂ݗL���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_SysCmd.bas/ExecDosCmd()
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͂����ꂼ��R�}���h���s()
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
' = �T�v    �I��͈͓��̌��������F��ύX����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͓��̌��������F��ύX()
    Const sMACRO_TITLE As String = "�I��͈͓��̌��������F��ύX"
    
    '�������F�ݒ聥����
    Const sCOLOR_TYPE As String = "0:�ԁA1:���A2:�΁A3:���A4:��A5:���A6:���A7:��"
    Const lCOLOR_NUM As Long = 8
    Dim vCOLOR_INFO() As Variant
    vCOLOR_INFO = _
        Array( _
            Array(255, 0, 0), _
            Array(75, 172, 198), _
            Array(118, 147, 60), _
            Array(112, 48, 160), _
            Array(247, 150, 70), _
            Array(255, 192, 0), _
            Array(255, 255, 255), _
            Array(0, 0, 0) _
        )
    '�������F�ݒ聣����
    
    Dim sSrchStr As String
    sSrchStr = InputBox("�������������͂��Ă�������", sMACRO_TITLE)
    
    Dim lColorIndex As Long
    lColorIndex = InputBox( _
        "�����F��I�����Ă�������" & vbNewLine & _
        "  " & sCOLOR_TYPE & vbNewLine _
        , _
        sMACRO_TITLE, _
        1 _
    )
    
    If lColorIndex < lCOLOR_NUM Then
        Dim oCell As Range
        For Each oCell In Selection
            Dim sTrgtStr As String
            sTrgtStr = oCell.Value
            Dim lStartIdx As Long
            lStartIdx = 1
            Do
                Dim lIdx As Long
                lIdx = InStr(lStartIdx, sTrgtStr, sSrchStr)
                If lIdx = 0 Then
                    Exit Do
                Else
                    lStartIdx = lIdx + Len(sSrchStr)
                    oCell.Characters(Start:=lIdx, Length:=Len(sSrchStr)).Font.Color = _
                        RGB( _
                            vCOLOR_INFO(lColorIndex)(0), _
                            vCOLOR_INFO(lColorIndex)(1), _
                            vCOLOR_INFO(lColorIndex)(2) _
                        )
                End If
            Loop While 1
        Next
        MsgBox "�����I", vbOKOnly, sMACRO_TITLE
    Else
        MsgBox "�����F�͎w��͈͓̔��őI�����Ă��������B" & vbNewLine & sCOLOR_TYPE, vbOKOnly, sMACRO_TITLE
    End If
End Sub

' =============================================================================
' = �T�v    �s���c���[�\���ɂ��ăO���[�v��
' =         Usage�F�c���[�O���[�v���������͈͂�I�����A�}�N���u�c���[���O���[�v���v�����s����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/TreeGroupSub()
' = ����    Macros.bas
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
' = �T�v    �V�[�g����ёւ���B
' =         �{���������s����ƁA�V�[�g���בւ���Ɨp�V�[�g���쐬����B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
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
            Set shWorkSht = .Sheets.Add(after:=.Sheets(.Sheets.Count))
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
' = �T�v    �V�[�g����ёւ���B
' =         �V�[�g���בւ���Ɨp�V�[�g�ɋL�ڂ̒ʂ�A�V�[�g����ёւ���B
' =         �K���V�[�g���בւ���Ɨp�V�[�g����Ăяo�����ƁI
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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
' = �T�v    �I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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
' = �T�v    �t�H���g�F���u�ԁv�́u�����v�Ńg�O������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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
' = �T�v    �w�i�F���u���v�́u�w�i�F�Ȃ��v�Ńg�O������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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
' = �T�v    �I�[�g�t�B�������s����B
' =         �w�肵�������ɉ����đI��͈͂��L���ăI�[�g�t�B�������s����B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
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

' =============================================================================
' = �T�v    �A�N�e�B�u�Z���R�����g�̂ݕ\�����A�ړ�����B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�( _
    ByVal sDirection As String _
)
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    '�A�N�e�B�u�Z���R�����g�\��
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    '�Z���ړ�
    Select Case sDirection
        Case "Right": ActiveCell.Offset(0, 1).Activate
        Case "Left": ActiveCell.Offset(0, -1).Activate
        Case "Down": ActiveCell.Offset(1, 0).Activate
        Case "Up": ActiveCell.Offset(-1, 0).Activate
        Case Else: Debug.Assert 1
    End Select
    
    '�A�N�e�B�u�Z���R�����g�\��
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�Z���R�����g�̂ݕ\������B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\��()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    '�A�N�e�B�u�Z���R�����g�\��
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    '�A�N�e�B�u�Z���R�����g�\��
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�Z���̃R�����g�\���̗L��/������؂�ւ���
' = �o��    �Ȃ�
' = �ˑ�    AddInSetting.cls/SearchWithKey()
' =         AddInSetting.cls/Update()
' =         AddInSetting.cls/Add()
' =         Macros.bas/DisableShortcutKeys()
' =         Macros.bas/UpdateShortcutKeySettings()
' =         Macros.bas/EnableShortcutKeys()
' = ����    Macros.bas
' =============================================================================
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\������шړ�_���[�h�ؑ�()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim bRet As Boolean
    Dim sValue As String
    bRet = clSetting.SearchWithKey(SETTING_KEY_CMNT_VSBL_ENB, sValue)
    If bRet = True Then
        If sValue = "True" Then
            MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ��𖳌��ɂ��܂�"
            Call DisableShortcutKeys
            Call clSetting.Update(SETTING_KEY_CMNT_VSBL_ENB, "False")
            Call UpdateShortcutKeySettings("Update")
            Call EnableShortcutKeys
        Else
            MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ���L���ɂ��܂�"
            Call DisableShortcutKeys
            Call clSetting.Update(SETTING_KEY_CMNT_VSBL_ENB, "True")
            Call UpdateShortcutKeySettings("Update")
            Call EnableShortcutKeys
        End If
        Debug.Assert bRet
    Else
        MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\������шړ��𖳌��ɂ��܂�"
        Call DisableShortcutKeys
        Call clSetting.Add(SETTING_KEY_CMNT_VSBL_ENB, "False")
        Call UpdateShortcutKeySettings("Update")
        Call EnableShortcutKeys
    End If
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�Z������n�C�p�[�����N��ɔ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �n�C�p�[�����N�Ŕ��()
    On Error Resume Next
    ActiveCell.Hyperlinks(1).Follow NewWindow:=True
    If Err.Number = 0 Then
        'Do Nothing
    Else
        Debug.Print "[" & Now & "] Error " & _
                    "[Macro] �n�C�p�[�����N�Ŕ�� " & _
                    "[Error No." & Err.Number & "] " & Err.Description
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N��MEMO�V�[�g�ֈړ�����
' = �o��    �Ȃ�
' = �ˑ��@�@Macros.bas/JumpToTrgtSheet()
' = ����    Macros.bas
' =============================================================================
Public Sub MEMO�V�[�g�փW�����v()
    Const TRGT_SHEET_NAME As String = "MEMO"
    Call JumpToTrgtSheet(TRGT_SHEET_NAME)
End Sub

' =============================================================================
' = �T�v    Excel���ᎆ
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub Excel���ᎆ()
    ActiveSheet.Cells.Select
    With Selection
        .ColumnWidth = 1.22
        .RowHeight = 10.8
        .Font.Size = 9
        .Font.Name = "�l�r �S�V�b�N"
    End With
    ActiveSheet.Cells(1, 1).Select
End Sub

' =============================================================================
' = �T�v    �񕝁A�s����������������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �����񕝒���()
    Selection.EntireColumn.AutoFit
End Sub
Public Sub �����s������()
    Selection.EntireRow.AutoFit
End Sub

' =============================================================================
' = �T�v    �őO�ʁA�Ŕw�ʂֈړ�����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �őO�ʂֈړ�()
    Selection.ShapeRange.ZOrder msoBringToFront
End Sub
Public Sub �Ŕw�ʂֈړ�()
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub

' =============================================================================
' = �T�v    EpTree�̊֐��c���[��Excel�Ŏ�荞��
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/ShowFilesSelectDialog()
' =         Mng_Collection.bas/ReadTxtFileToCollection()
' =         Mng_String.bas/ExecRegExp()
' =         Mng_ExcelOpe.bas/CreateNewWorksheet()
' = ����    Macros.bas
' =============================================================================
Public Sub EpTree�̊֐��c���[��Excel�Ŏ�荞��()
    Const STRT_ROW As Long = 2
    Const STRT_CLM As Long = 2
    Const SHEET_NAME As String = "�֐��c���["
    Const MAX_FUNC_LEVEL_INI As Long = 10
    Const CLM_WIDTH As Long = 2
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    '�t�@�C���p�X����
    Dim sTrgtFilePath As String
    Dim asSelectedFiles() As String
    Call ShowFilesSelectDialog(asSelectedFiles)
    sTrgtFilePath = asSelectedFiles(0)
    
    '�V�[�g�ǉ�
    Dim sSheetName As String
    Dim shTrgtSht As Worksheet
    sSheetName = CreateNewWorksheet(SHEET_NAME)
    Set shTrgtSht = ActiveWorkbook.Sheets(sSheetName)
    
    '�e�L�X�g�t�@�C���ǂݏo��
    Dim cFileContents As Collection
    Set cFileContents = New Collection
    Call ReadTxtFileToCollection(sTrgtFilePath, cFileContents)
    
    '�t�@�C���c���[�o��
    lStrtRow = STRT_ROW
    lStrtClm = STRT_CLM
    lRowIdx = lStrtRow
    
    With shTrgtSht
        .Cells(lRowIdx, lStrtClm + 0).Value = "�t�@�C���p�X"
        .Cells(lRowIdx, lStrtClm + 1).Value = "�s��"
        .Cells(lRowIdx, lStrtClm + 2).Value = "�֐���"
        .Cells(lRowIdx, lStrtClm + 3).Value = "�֐��c���["
    End With
    lRowIdx = lRowIdx + 1
    
    Dim lMaxFuncLevel As Long
    lMaxFuncLevel = MAX_FUNC_LEVEL_INI
    Dim vFileLine As Variant
    For Each vFileLine In cFileContents
        Dim oMatchResult As Object
        Call ExecRegExp( _
            vFileLine, _
            "^(.+)? +(\d+): (  )?([��|��|��|  ]*)(\w+)(��)?", _
            oMatchResult _
        )
        
        Dim sFilePath As String
        Dim sLineNo As String
        Dim lFuncLevel As Long
        Dim sFuncName As String
        Dim sOmission As String
        sFilePath = oMatchResult(0).SubMatches(0)
        sLineNo = oMatchResult(0).SubMatches(1)
        If sLineNo = 0 Then
            sLineNo = ""
        End If
        lFuncLevel = LenB(oMatchResult(0).SubMatches(3)) / 2
        sFuncName = oMatchResult(0).SubMatches(4)
        sOmission = String(LenB(oMatchResult(0).SubMatches(5)) / 2, "��")
        
        With shTrgtSht
            .Cells(lRowIdx, lStrtClm + 0).Value = sFilePath
            .Cells(lRowIdx, lStrtClm + 1).Value = sLineNo
            .Cells(lRowIdx, lStrtClm + 2).Value = sFuncName
            .Cells(lRowIdx, lStrtClm + 3 + lFuncLevel).Value = sFuncName & sOmission
        End With
        If lFuncLevel > lMaxFuncLevel Then
            lMaxFuncLevel = lFuncLevel
        End If
        
        lRowIdx = lRowIdx + 1
    Next
    
    With shTrgtSht
        lLastClm = STRT_CLM + 3 + lMaxFuncLevel
        lLastRow = lRowIdx
        
        '�^�C�g���s ��������
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lStrtRow, lStrtClm + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lStrtRow, lLastClm)).HorizontalAlignment = xlCenterAcrossSelection
        
        '�񕝒���
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lLastRow, lStrtClm + 0)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 1), .Cells(lLastRow, lStrtClm + 1)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 2), .Cells(lLastRow, lStrtClm + 2)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lLastRow, lLastClm)).ColumnWidth = CLM_WIDTH
    
        '�I�[�g�t�B���^
        .Range(.Cells(lStrtRow, lStrtClm), .Cells(lLastRow, lLastClm)).AutoFilter
        
        '�s����
        .Rows(lStrtRow).RowHeight = .Rows(lStrtRow).RowHeight * 3
        
        '�^�C�g����Œ�
        ActiveWindow.FreezePanes = False
        .Rows(lStrtRow + 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "�֐��c���[�쐬�����I"
End Sub

' *****************************************************************************
' * �����p�}�N��
' *****************************************************************************
' =============================================================================
' = �T�v    �ݒ荀�ڈꗗ���o��
' = �o��    �Ȃ�
' = �ˑ�    AddInSetting.cls/SearchWithIdx()
' = ����    Macros.bas
' =============================================================================
Private Sub OutputSettingList()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lSettingNum As Long
    lSettingNum = clSetting.Count
    
    Debug.Print ""
    Debug.Print "*** �ݒ荀�ڈꗗ�o�� ***"
    If lSettingNum = 0 Then
        'Do Nothing
    Else
        Dim lSettingIdx As Long
        For lSettingIdx = 1 To lSettingNum
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lSettingIdx, sSettingKey, sSettingValue)
            Debug.Print sSettingKey & " = " & sSettingValue
        Next lSettingIdx
    End If
End Sub

' =============================================================================
' = �T�v    ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ��ǉ�
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/UpdateShortcutKeySettings()
' = ����    Macros.bas
' =============================================================================
Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ��ǉ�()
    Call UpdateShortcutKeySettings("Add")
End Sub

' =============================================================================
' = �T�v    ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ���폜
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/UpdateShortcutKeySettings()
' = ����    Macros.bas
' =============================================================================
Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ���폜()
    Call UpdateShortcutKeySettings("Delete")
End Sub

' =============================================================================
' = �T�v    ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ���X�V
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/UpdateShortcutKeySettings()
' = ����    Macros.bas
' =============================================================================
Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[�ݒ���X�V()
    Call UpdateShortcutKeySettings("Update")
End Sub

' =============================================================================
' = �T�v    ���[�U�[��`�V���[�g�J�b�g�L�[��L����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/EnableShortcutKeys()
' = ����    Macros.bas
' =============================================================================
Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[��L����()
    Call EnableShortcutKeys
End Sub

' =============================================================================
' = �T�v    ���[�U�[��`�V���[�g�J�b�g�L�[�𖳌���
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/DisableShortcutKeys()
' = ����    Macros.bas
' =============================================================================
Public Sub ���[�U�[��`�V���[�g�J�b�g�L�[�𖳌���()
    Call DisableShortcutKeys
End Sub

' *****************************************************************************
' * �����֐���`
' *****************************************************************************
' ==================================================================
' = �T�v    ���� �^�ϊ�(String��Long)
' = ����    sNum            String  [in]  ����(String�^)
' = �ߒl                    Long          ����(Long�^)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function

' ==================================================================
' = �T�v    ���� �^�ϊ�(Long��String)
' = ����    lNum            Long    [in]    ����(Long�^)
' = �ߒl                    String          ����(String�^)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function

' ==================================================================
' = �T�v    �c���[���O���[�v��
' = ����    shTrgtSht       Worksheet   [in,out]    ���[�N�V�[�g
' = ����    lGrpStrtRow     Long        [in]        �擪�s
' = ����    lGrpLastRow     Long        [in]        �����s
' = ����    lGrpStrtClm     Long        [in]        �擪��
' = ����    lGrpLastClm     Long        [in]        ������
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/IsGroupParent()
' = ����    Macros.bas
' ==================================================================
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

' ==================================================================
' = �T�v    �w�肵���Z���̒����Z�����󔒂ŁA�E���Z�����󔒂łȂ��ꍇ�A
' =         �O���[�v�̐e�ł���Ɣ��f����B
' = ����    shTrgtSht   Worksheet   [in,out]    ���[�N�V�[�g
' = ����    lRow        Long        [in]        �s
' = ����    lClm        Long        [in]        ��
' = �ߒl                Boolean
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
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
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
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

' ==================================================================
' = �T�v    �R�}���h�����s
' = ����    sCommand    String   [in]   �R�}���h
' = �ߒl                String          �W���o��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_SysCmd.bas
' ==================================================================
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

' ==================================================================
' = �T�v    �V���[�g�J�b�g�L�[�ݒ��ǉ�/�폜
' = ����    sKey        String   [in]   �L�[
' = ����    sMacroName  String   [in]   �}�N����
' = ����    sMode       String   [in]   ���[�h(Add/Update/Delete)
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    AddInSetting.cls/Add()
' =         AddInSetting.cls/Update()
' =         AddInSetting.cls/Delete()
' = ����    Macros.bas
' ==================================================================
Private Function UpdateShtcutSetting( _
    ByVal sKey As String, _
    ByVal sMacroName As String, _
    ByVal sMode As String _
)
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim sSettingKey As String
    Dim sSettingValue As String
    sSettingKey = SHTCUTKEY_KEYWORD_PREFIX & "_" & sMacroName
    sSettingValue = sKey
    Select Case sMode
        Case "Add": Call clSetting.Add(sSettingKey, sSettingValue)
        Case "Update": Call clSetting.Update(sSettingKey, sSettingValue)
        Case "Delete": Call clSetting.Delete(sSettingKey)
        Case Else: Debug.Assert False
    End Select
End Function

' ==================================================================
' = �T�v    �V���[�g�J�b�g�L�[��L����
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    AddInSetting.cls/Count()
' =         AddInSetting.cls/SearchWithIdx()
' = ����    Macros.bas
' ==================================================================
Private Function EnableShortcutKeys()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lNum As Long
    lNum = clSetting.Count
    If lNum = 0 Then
        'Do Nothing
    Else
        Dim lLastRow As Long
        lLastRow = lNum
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lRowIdx, sSettingKey, sSettingValue)
            '*** �V���[�g�J�b�g�ݒ�̏ꍇ ***
            If InStr(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX) Then
                Dim sShrcutMacroName As String
                Dim sShtcutKey As String
                sShrcutMacroName = Replace(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX & "_", "")
                sShtcutKey = sSettingValue
                If sShtcutKey = "" Then
                    'Do Nothing
                Else
                    Application.OnKey sShtcutKey, sShrcutMacroName
                End If
            
            '*** �V���[�g�J�b�g�ݒ�łȂ��ꍇ ***
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End If
End Function

' ==================================================================
' = �T�v    �V���[�g�J�b�g�L�[�𖳌���
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    AddInSetting.cls/Count()
' =         AddInSetting.cls/SearchWithIdx()
' = ����    Macros.bas
' ==================================================================
Private Function DisableShortcutKeys()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lNum As Long
    lNum = clSetting.Count
    If lNum = 0 Then
        'Do Nothing
    Else
        Dim lLastRow As Long
        lLastRow = lNum
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lRowIdx, sSettingKey, sSettingValue)
            If InStr(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX) Then
                Dim sShtcutKey As String
                sShtcutKey = sSettingValue
                If sShtcutKey = "" Then
                    'Do Nothing
                Else
                    Application.OnKey sShtcutKey
                End If
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End If
End Function

' ============================================
' = �T�v    �z��̓��e���t�@�C���ɏ������ށB
' = ����    sFilePath     String  [in]  �o�͂���t�@�C���p�X
' =         asFileLine()  String  [in]  �o�͂���t�@�C���̓��e
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ============================================
Private Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String = "shift_jis" _
)
    Dim oTxtObj As Object
    Dim lLineIdx As Long
    
    If Sgn(asFileLine) = 0 Then
        'Do Nothing
    Else
        Set oTxtObj = CreateObject("ADODB.Stream")
        With oTxtObj
            .Type = 2
            .Charset = sCharSet
            .Open
            
            '�z���1�s���I�u�W�F�N�g�ɏ�������
            For lLineIdx = 0 To UBound(asFileLine)
                .WriteText asFileLine(lLineIdx), 1
            Next lLineIdx
            
            .SaveToFile (sFilePath), 2    '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
            .Close
        End With
    End If
    
    Set oTxtObj = Nothing
End Function

' ==================================================================
' = �T�v    �t�@�C���i�����j�I���_�C�A���O��\������
' = ����    asSelectedFiles String()    [out] �I�����ꂽ�t�@�C���p�X�ꗗ
' = ����    sInitPath       String      [in]  �f�t�H���g�t�@�C���p�X�i�ȗ��j
' = ����    sFilters        String      [in]  �I�����̃t�B���^�i�ȗ��j(��)
' = �ߒl    �Ȃ�
' = �o��    (��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =              ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =                    �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =                    �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =                    �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' = �ˑ�    Mng_FileSys.bas/SetDialogFilters()
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function ShowFilesSelectDialog( _
    ByRef asSelectedFiles() As String, _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
)
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "�t�@�C����I�����Ă��������i�����j"
    fdDialog.AllowMultiSelect = True
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
 
    '�_�C�A���O�\��
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ReDim Preserve asSelectedFiles(0)
        asSelectedFiles(0) = ""
    Else
        Dim lSelNum As Long
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asSelectedFiles(lSelNum - 1)
        Dim lSelIdx As Long
        For lSelIdx = 0 To lSelNum - 1
            Dim sSelectedPath As String
            sSelectedPath = fdDialog.SelectedItems(lSelIdx + 1)
            If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
                asSelectedFiles(lSelIdx) = sSelectedPath
            Else
                asSelectedFiles(lSelIdx) = ""
            End If
        Next lSelIdx
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFilesSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
        sFilters = "�S�Ẵt�@�C��/*.*,�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
 
        Dim asSelectedFiles() As String
        Call ShowFilesSelectDialog( _
                    asSelectedFiles, _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
        Dim sBuf As String
        sBuf = ""
        sBuf = sBuf & vbNewLine & UBound(asSelectedFiles) + 1
        Dim lSelIdx As Long
        For lSelIdx = 0 To UBound(asSelectedFiles)
            sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
        Next lSelIdx
        MsgBox sBuf
    End Sub
 
' ==================================================================
' = �T�v    ShowFileSelectDialog() �� ShowFilesSelectDialog() �p�̊֐�
' =         �_�C�A���O�̃t�B���^��ǉ�����B�w����@�͈ȉ��B
' =           ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =               �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =               �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =               �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���󕶎��̏ꍇ�A�t�B���^���N���A����B
' = ����    sFilters    String  [in]    �t�B���^
' = ����    fdDialog    String  [out]   �_�C�A���O
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                            """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "�����𒆒f���܂��B"
                        End
                    End If
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Else
                MsgBox _
                    "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                    "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                    "  " & sFilters
                MsgBox "�����𒆒f���܂��B"
                End
            End If
        End If
    End If
End Function

' ==================================================================
' = �T�v    ���[�N�V�[�g��V�K�쐬
' =         �d���������[�N�V�[�g������ꍇ�A_1, _2 ...�ƘA�ԂɂȂ�B
' =         �Ăяo�����ɂ͍쐬�������[�N�V�[�g����Ԃ��B
' = ����    sSheetName  String  [in]    �V�[�g��
' = �ߒl                                �V�[�g��
' = �o��    �Ȃ�
' = �ˑ�    Mng_ExcelOpe.bas/ExistsWorksheet()
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Private Function CreateNewWorksheet( _
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
' = ����    sTrgtShtName    String  [in]    �V�[�g��
' = �ߒl                                    ���݃`�F�b�N����
' = �o��    �Ȃ�
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
' = �T�v    �e�L�X�g�t�@�C���̒��g��z��Ɋi�[
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [out]   �t�@�C���̒��g
' = �ߒl    �ǂݏo������    Boolean             �ǂݏo������
' =                                                 True:�t�@�C������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Collection.bas
' ==================================================================
Private Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath As String, _
    ByRef cFileContents As Collection _
)
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile As Object
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.Add objTxtFile.ReadLine
            Loop
            ReadTxtFileToCollection = True
        Else
            ReadTxtFileToCollection = False
        '   WScript.Echo "�G���[ " & Err.Description
        End If
        
        objTxtFile.Close
    Else
        ReadTxtFileToCollection = False
    End If
    On Error GoTo 0
End Function
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileContents As Collection
        Set cFileContents = New Collection
        Dim sInFilePath As String
        sInFilePath = "C:\codes\vbs\_lib\Test.csv"
        Dim bRet As Boolean
        bRet = ReadTxtFileToCollection(sInFilePath, cFileContents)
    End Sub

' ==================================================================
' = �T�v    ���K�\���������s���iVba�}�N���֐��p�j
' = ����    sTargetStr      String  [in]  �����Ώە�����
' = ����    sSearchPattern  String  [in]  �����p�^�[��
' = ����    oMatchResult    Object  [out] ��������
' = �ߒl                    Boolean       �q�b�g�L��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function ExecRegExp( _
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
    ByRef oMatchResult As Object, _
    Optional ByVal bIgnoreCase As Boolean = True, _
    Optional ByVal bGlobal As Boolean = True _
) As Boolean
    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.IgnoreCase = bIgnoreCase
    oRegExp.Global = bGlobal
    oRegExp.Pattern = sSearchPattern
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count = 0 Then
        ExecRegExp = False
    Else
        ExecRegExp = True
    End If
End Function
    Private Sub Test_ExecRegExp()
        Dim sTargetStr As String
        Dim oMatchResult As Object
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print ExecRegExp(sTargetStr, " \w+\(", oMatchResult)
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �N���b�v�{�[�h�Ƀe�L�X�g���R�s�[�iWin32Api���g�p�j
' = ����    sText       String  [in]  �R�s�[�Ώە�����
' = �ߒl                Boolean       �R�s�[����
' = �o��    Win32API���g�p����B
' =         �� �N���b�v�{�[�h�� DataObject �� PutInClipboard �ł����p
' =            �\��������DataObject �͎Q�Ɛݒ肪�K�v�Ȃ��������̃N
' =            ���b�v�{�[�h�`���ɂ͓\��t������Ȃ���iCF_UNICODETEXT
' =            �݂̂� CF_TEXT�ւ͓\��t������Ȃ��j
' =            ��L�̂悤�� DataObject ���g�p�������Ȃ��ꍇ�ɖ{�֐�
' =            �𗘗p���邱�ơ
' = �ˑ�    user32/OpenClipboard()
' =         user32/EmptyClipboard()
' =         user32/CloseClipboard()
' =         user32/SetClipboardData()
' =         kernel32/GlobalAlloc()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = ����    Mng_Clipboard.bas
' ==================================================================
Public Function SetToClipboard( _
    sText As String _
) As Boolean
    '�萔�錾
    Const GMEM_MOVEABLE         As Long = &H2
    Const GMEM_ZEROINIT         As Long = &H40
    Const GHND                  As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
    Const CF_TEXT               As Long = 1
    Const CF_OEMTEXT            As Long = 7
    
    Dim hGlobal As Long
    Dim lTextLen As Long
    Dim p As Long
    
    '�߂�l���Ƃ肠�����AFalse�ɐݒ肵�Ă����B
    If OpenClipboard(0) <> 0 Then
        If EmptyClipboard() <> 0 Then
            lTextLen = LenB(sText) + 1 '�����̎Z�o(�{����Unicode����ϊ���̒������g���ق����悢)
            hGlobal = GlobalAlloc(GHND, lTextLen) '�R�s�[��̗̈�m��
            p = GlobalLock(hGlobal)
            Call lstrcpy(p, sText) '��������R�s�[
            Call GlobalUnlock(hGlobal) '�N���b�v�{�[�h�ɓn���Ƃ��ɂ�Unlock���Ă����K�v������
            Call SetClipboardData(CF_TEXT, hGlobal) '�N���b�v�{�[�h�֓\��t����
            Call CloseClipboard '�N���b�v�{�[�h���N���[�Y
            SetToClipboard = True '�R�s�[����
        Else
            SetToClipboard = False
        End If
    Else
        SetToClipboard = False
    End If
End Function

' ==================================================================
' = �T�v    �A�N�e�B�u�u�b�N�̎w��V�[�g�ֈړ�����
' = ����    sSheetName      String  [in]  �ړ��ΏۃV�[�g��
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Public Function JumpToTrgtSheet( _
    ByVal sSheetName As String _
)
    Dim shSheet As Worksheet
    For Each shSheet In ActiveWorkbook.Sheets
        If shSheet.Name = sSheetName Then
            shSheet.Activate
        End If
    Next
End Function


