Attribute VB_Name = "Macros"
Option Explicit

' my excel addin macros v2.51

' =============================================================================
' =  <<�}�N���ꗗ>>
' =     �E����
' =         F1�w���v������                              F1�w���v�𖳌�������
' =
' =     �E�}�N���ݒ�
' =         �}�N���V���[�g�J�b�g�L�[�S�ėL����          �}�N���V���[�g�J�b�g�L�[�S�ėL����
' =         �}�N���V���[�g�J�b�g�L�[�S�Ė�����          �}�N���V���[�g�J�b�g�L�[�S�Ė�����
' =         �A�h�C���}�N�����s                          �A�h�C���}�N�����s
' =
' =     �E�V�[�g����
' =         EpTree�̊֐��c���[��Excel�Ŏ�荞��         EpTree�̊֐��c���[��Excel�Ŏ�荞��
' =         Excel���ᎆ                                 Excel���ᎆ
' =         �I���V�[�g�؂�o��                          �I���V�[�g��ʃt�@�C���ɐ؂�o��
' =         �S�V�[�g�����R�s�[                          �u�b�N���̃V�[�g����S�ăR�s�[����
' =         �V�[�g�\����\����؂�ւ�                  �V�[�g�\��/��\����؂�ւ���
' =         �V�[�g���בւ���Ɨp�V�[�g���쐬            �V�[�g���בւ���Ɨp�V�[�g�쐬
' =         �V�[�g�I���E�B���h�E��\��                  �V�[�g�I���E�B���h�E��\������
' =         �擪�V�[�g�փW�����v                        �A�N�e�B�u�u�b�N�̐擪�V�[�g�ֈړ�����
' =         �����V�[�g�փW�����v                        �A�N�e�B�u�u�b�N�̖����V�[�g�ֈړ�����
' =
' =     �E�Z������
' =         �t�@�C���G�N�X�|�[�g                        �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =         DOS�R�}���h���ꊇ���s                       �I��͈͓���DOS�R�}���h���܂Ƃ߂Ď��s����B
' =         DOS�R�}���h���e�X���s                       �I��͈͓���DOS�R�}���h�����ꂼ����s����B
' =         ���������̕����F��ύX                      �I��͈͓��̌��������̕����F��ύX����
' =         �Z�����̊ې������f�N�������g                �A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' =         �Z�����̊ې������C���N�������g              �@�`�M���w�肵�āA�w��ԍ��ȍ~���f�N�������g����
' =         �c���[���O���[�v��                          �c���[�O���[�v������
' =         �n�C�p�[�����N�ꊇ�I�[�v��                  �I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' =         �n�C�p�[�����N�Ŕ��                        �A�N�e�B�u�Z������n�C�p�[�����N��ɔ��
' =         �I��͈͓��Œ���                            �I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' =         �͈͂��ێ������܂܃Z���R�s�[                �I��͈͂�͈͂��ێ������܂܃Z���R�s�[����B(�_�u���N�I�[�e�[�V����������)
' =         ��s�ɂ܂Ƃ߂ăZ���R�s�[                    �I��͈͂���s�ɂ܂Ƃ߂ăZ���R�s�[����B
' =         ���ݒ�ύX����s�ɂ܂Ƃ߂ăZ���R�s�[        ��s�ɂ܂Ƃ߂ăZ���R�s�[�ɂĎg�p����u�擪����,��؂蕶��,���������v��ύX����
' =         �t�H���g�F���g�O��                          �t�H���g�F���u�ݒ�F�v�́u�����v�Ńg�O������
' =         �w�i�F���g�O��                              �w�i�F���u�ݒ�F�v�́u�w�i�F�Ȃ��v�Ńg�O������
' =         ���ݒ�ύX���t�H���g�F���g�O���̐F�I��      �u�t�H���g�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' =         ���ݒ�ύX���w�i�F���g�O���̐F�I��          �u�w�i�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' =         ���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g  �u�t�H���g�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' =         ���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g      �u�w�i�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' =         �I�[�g�t�B�����s                            �I�[�g�t�B�������s����
' =         �C���f���g���グ��                          �C���f���g���グ��
' =         �C���f���g��������                          �C���f���g��������
' =         ���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��  �A�N�e�B�u�Z���R�����g�ݒ��؂�ւ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\��              ���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉ��ړ�    ���ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ď�ړ�    ��ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉE�ړ�    �E�ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����č��ړ�    ���ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         Excel�������`�����{                         Excel�������`�����{
' =         Excel�������`������                         Excel�������`������
' =
' =     �E�I�u�W�F�N�g����
' =         �őO�ʂֈړ�                                �őO�ʂֈړ�����
' =         �Ŕw�ʂֈړ�                                �Ŕw�ʂֈړ�����
' =============================================================================

'******************************************************************************
'* ���O����
'******************************************************************************
'Win32API�錾
'������Macro.bas/�͈͂��ێ������܂܃Z���R�s�[()������
'������Macro.bas/��s�ɂ܂Ƃ߂ăZ���R�s�[()������
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'������Macro.bas/��s�ɂ܂Ƃ߂ăZ���R�s�[()������
'������Macro.bas/�͈͂��ێ������܂܃Z���R�s�[()������

'������Mng_Clipboard.bas/SetToClipboard()������
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

'������Macro.bas/ShowColorPalette()������
Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'������Macro.bas/ShowColorPalette()������

'******************************************************************************
'* �ݒ�l
'******************************************************************************
'������ �ݒ�(�����l) ������
'=== �w�i�F���g�O��()/�t�H���g�F���g�O��() ===
    '[�F���Q�l] https://excel-toshokan.com/vba-color-list/
    Const lCLRTGLBG_CLR_RGB As Long = vbYellow
    Const lCLRTGLFONT_CLR_RGB As Long = vbRed
'=== �A�N�e�B�u�Z���R�����g�ݒ�() ===
    Const bCMNT_VSBL_ENB As Boolean = False
'=== Excel���ᎆ() ===
    Const sEXCELGRID_FONT_NAME As String = "�l�r �S�V�b�N"
    Const lEXCELGRID_FONT_SIZE As Long = 9
    Const lEXCELGRID_CLM_WIDTH As Long = 3 '3������
'=== ���������̕����F��ύX() ===
    Const sWORDCOLOR_SRCH_WORD As String = ""
    Const lWORDCOLOR_CLR_RGB As Long = vbRed
'=== �t�@�C���G�N�X�|�[�g() ===
    Const sFILEEXPORT_OUT_FILE_NAME As String = "MyExcelAddinFileExport.csv"
    Const bFILEEXPORT_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOS�R�}���h���ꊇ���s() ===
    Const sCMDEXEBAT_BAT_FILE_NAME As String = "MyExcelAddinCmdexebat.bat"
    Const sCMDEXEBAT_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdexebat.log"
    Const bCMDEXEBAT_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOS�R�}���h���e�X���s() ===
    Const sCMDEXEUNI_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdexeuni.log"
    Const bCMDEXEUNI_IGNORE_INVISIBLE_CELL As Boolean = True
'=== EpTree�̊֐��c���[��Excel�Ŏ�荞��() ===
    Const sEPTREE_OUT_SHEET_NAME As String = "CallTree"
    Const lEPTREE_MAX_FUNC_LEVEL_INI As Long = 10
    Const lEPTREE_CLM_WIDTH As Long = 2
    Const sEPTREE_OUT_LOG_PATH As String = "c:\"
    Const sEPTREE_DEV_ROOT_DIR_PATH As String = "c:\"
    Const lEPTREE_DEV_ROOT_DIR_LEVEL As Long = 0
'=== �͈͂��ێ������܂܃Z���R�s�[() ===
    Const bCELLCOPYRNG_IGNORE_INVISIBLE_CELL As Boolean = True
    Const sCELLCOPYRNG_DELIMITER As String = vbTab
'=== ��s�ɂ܂Ƃ߂ăZ���R�s�[() ===
    Const bCELLCOPYLINE_IGNORE_INVISIBLE_CELL As Boolean = True
    Const bCELLCOPYLINE_IGNORE_BLANK_CELL As Boolean = True
    Const sCELLCOPYLINE_PREFFIX As String = "("
    Const sCELLCOPYLINE_DELIMITER As String = "|"
    Const sCELLCOPYLINE_SUFFIX As String = ")"
'=== �V�[�g�I���E�B���h�E��\��() ===
    Const bSHTSELWIN_MSGBOX_SHOW As Boolean = False
'������ �ݒ� ������

' ==================================================================
' = �T�v    �V���[�g�J�b�g�L�[�̗L��/������؂�ւ���
' = ����    bActivateShortcutKeys   Boolean     [in]    �L����/������
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' ==================================================================
Private Sub SwitchMacroShortcutKeysActivation( _
    ByVal bActivateShortcutKeys As Boolean _
)
    Dim dMacroShortcutKeys As Object
    Set dMacroShortcutKeys = CreateObject("Scripting.Dictionary")
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim bCmntVsblEnb As Boolean
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCMNT_VSBL_ENB", bCmntVsblEnb, bCMNT_VSBL_ENB, False)
    
    '*** �V���[�g�J�b�g�L�[�ݒ�X�V ***
    ' <<�V���[�g�J�b�g�L�[�ǉ����@>>
    '   dMacroShortcutKeys�ɑ΂��ăL�[<�}�N����>�A�l<�V���[�g�J�b�g�L�[>��ǉ�����B
    '   �������ɂ̓V���[�g�J�b�g�L�[�A�������Ƀ}�N�������w�肷��B
    '   �V���[�g�J�b�g�L�[�� Ctrl �� Shift �ȂǂƑg�ݍ��킹�Ďw��ł���B
    '     Ctrl�F^�AShift�F+�AAlt�F%
    '   �ڍׂ͈ȉ� URL �Q�ƁB
    '     https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '������ �ݒ� ������
'   dMacroShortcutKeys.Add "", "�I��͈͓��Œ���"
    
    dMacroShortcutKeys.Add "^+c", "�͈͂��ێ������܂܃Z���R�s�["
    dMacroShortcutKeys.Add "^+d", "��s�ɂ܂Ƃ߂ăZ���R�s�["
    dMacroShortcutKeys.Add "^%d", "���ݒ�ύX����s�ɂ܂Ƃ߂ăZ���R�s�["
    dMacroShortcutKeys.Add "+%{F8}", "�A�h�C���}�N�����s"
    
'   dMacroShortcutKeys.Add "", "�t�@�C���G�N�X�|�[�g"
'   dMacroShortcutKeys.Add "", "DOS�R�}���h���e�X���s"
'   dMacroShortcutKeys.Add "", "DOS�R�}���h���ꊇ���s"
'   dMacroShortcutKeys.Add "^+f", "���������̕����F��ύX" '�uCtrlShiftF�}�N���v�ɂĎ��s
    
    dMacroShortcutKeys.Add "^%c", "�S�V�[�g�����R�s�["
'   dMacroShortcutKeys.Add "", "�I���V�[�g�؂�o��"
'   dMacroShortcutKeys.Add "", "�V�[�g�\����\����؂�ւ�"
'   dMacroShortcutKeys.Add "", "�V�[�g���בւ���Ɨp�V�[�g���쐬"
    dMacroShortcutKeys.Add "^%{PGUP}", "�V�[�g�I���E�B���h�E��\��"
    dMacroShortcutKeys.Add "^%{PGDN}", "�V�[�g�I���E�B���h�E��\��"
    dMacroShortcutKeys.Add "^%{HOME}", "�擪�V�[�g�փW�����v"
    dMacroShortcutKeys.Add "^%{END}", "�����V�[�g�փW�����v"
    
'   dMacroShortcutKeys.Add "", "�Z�����̊ې������f�N�������g"
'   dMacroShortcutKeys.Add "", "�Z�����̊ې������C���N�������g"
    
'   dMacroShortcutKeys.Add "", "�c���[���O���[�v��"
'   dMacroShortcutKeys.Add "", "�n�C�p�[�����N�ꊇ�I�[�v��"
    
    dMacroShortcutKeys.Add "^2", "�w�i�F���g�O��"
    dMacroShortcutKeys.Add "^%2", "���ݒ�ύX���w�i�F���g�O���̐F�I��"
    dMacroShortcutKeys.Add "+%2", "���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g"
    dMacroShortcutKeys.Add "^3", "�t�H���g�F���g�O��"
    dMacroShortcutKeys.Add "^%3", "���ݒ�ύX���t�H���g�F���g�O���̐F�I��"
    dMacroShortcutKeys.Add "+%3", "���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g"

    dMacroShortcutKeys.Add "^%{DOWN}", "'�I�[�g�t�B�����s(""Down"")'"
    dMacroShortcutKeys.Add "^%{UP}", "'�I�[�g�t�B�����s(""Up"")'"
    
    dMacroShortcutKeys.Add "^%{RIGHT}", "�C���f���g���グ��"
    dMacroShortcutKeys.Add "^%{LEFT}", "�C���f���g��������"
    
    dMacroShortcutKeys.Add "^+{F11}", "���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��"
    dMacroShortcutKeys.Add "^+j", "�n�C�p�[�����N�Ŕ��"
    
    dMacroShortcutKeys.Add "^%h", "Excel���ᎆ"
'   dMacroShortcutKeys.Add "", "EpTree�̊֐��c���[��Excel�Ŏ�荞��"
    
'   dMacroShortcutKeys.Add "", "�����񕝒���"
'   dMacroShortcutKeys.Add "", "�����s������"
    
    dMacroShortcutKeys.Add "^+f", "CtrlShiftF�}�N��"
'   dMacroShortcutKeys.Add "^+f", "�őO�ʂֈړ�" '�uCtrlShiftF�}�N���v�ɂĎ��s
    dMacroShortcutKeys.Add "^+b", "�Ŕw�ʂֈړ�"
    
    dMacroShortcutKeys.Add "^+i", "Excel�������`�����{"
    dMacroShortcutKeys.Add "^%i", "Excel�������`������"
    
    If bCmntVsblEnb = True Then
        dMacroShortcutKeys.Add "{DOWN}", "�A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉ��ړ�"
        dMacroShortcutKeys.Add "{UP}", "�A�N�e�B�u�Z���R�����g�̂ݕ\�����ď�ړ�"
        dMacroShortcutKeys.Add "{RIGHT}", "�A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉE�ړ�"
        dMacroShortcutKeys.Add "{LEFT}", "�A�N�e�B�u�Z���R�����g�̂ݕ\�����č��ړ�"
    Else
        dMacroShortcutKeys.Add "{DOWN}", ""
        dMacroShortcutKeys.Add "{UP}", ""
        dMacroShortcutKeys.Add "{RIGHT}", ""
        dMacroShortcutKeys.Add "{LEFT}", ""
    End If
    '������ �ݒ� ������
    
    '*** �V���[�g�J�b�g�L�[�ݒ蔽�f ***
    Dim vShortcutKey As Variant
    Dim sMacroName As String
    If bActivateShortcutKeys = True Then
        For Each vShortcutKey In dMacroShortcutKeys
            sMacroName = dMacroShortcutKeys.Item(vShortcutKey)
            If sMacroName = "" Then
                Application.OnKey CStr(vShortcutKey)              '�V���[�g�J�b�g�L�[�N���A
            Else
                Application.OnKey CStr(vShortcutKey), sMacroName  '�V���[�g�J�b�g�L�[�ݒ�
            End If
        Next
    Else
        For Each vShortcutKey In dMacroShortcutKeys
            Application.OnKey CStr(vShortcutKey)                  '�V���[�g�J�b�g�L�[�N���A
        Next
    End If
End Sub

' *****************************************************************************
' * �O�����J�p�}�N��
' *****************************************************************************
' =============================================================================
' = �T�v    �}�N���V���[�g�J�b�g�L�[�S�ėL����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub �}�N���V���[�g�J�b�g�L�[�S�ėL����()
    Call SwitchMacroShortcutKeysActivation(True)
End Sub

' =============================================================================
' = �T�v    �}�N���V���[�g�J�b�g�L�[�S�Ė�����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub �}�N���V���[�g�J�b�g�L�[�S�Ė�����()
    Call SwitchMacroShortcutKeysActivation(False)
End Sub

' =============================================================================
' = �T�v    �V���[�g�J�b�g�L�[�d�����̐U�蕪�������iCtrl+Shift+F�j
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/�őO�ʂֈړ�()
' =         Macros.bas/���������̕����F��ύX()
' = ����    Macros.bas
' =============================================================================
Public Sub CtrlShiftF�}�N��()
    On Error Resume Next
    Call �őO�ʂֈړ�
    If Err.Number <> 0 Then
        Call ���������̕����F��ύX
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' = �T�v    F1�w���v�𖳌�������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub F1�w���v������()
    Application.OnKey "{F1}", ""
End Sub

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
    Const sMACRO_NAME As String = "�Z�����̊ې������f�N�������g"
    Const NUM_MAX As Long = 15
    Const NUM_MIN As Long = 1
    
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
    MsgBox "�u�������I", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = �T�v    �A�`�N���w�肵�āA�w��ԍ��ȍ~���C���N�������g����
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = ����    Macros.bas
' =============================================================================
Public Sub �Z�����̊ې������C���N�������g()
    Const sMACRO_NAME As String = "�Z�����̊ې������C���N�������g"
    Const NUM_MAX As Long = 15
    Const NUM_MIN As Long = 1
    
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
    MsgBox "�u�������I", vbOKOnly, sMACRO_NAME
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
    Const sMACRO_NAME As String = "�S�V�[�g�����R�s�["
    
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
    
    MsgBox "�u�b�N���̃V�[�g����S�ăR�s�[���܂���", vbOKOnly, sMACRO_NAME
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
' = �T�v    �V�[�g�I���E�B���h�E��\������
' = �o��    �E���ɂ���Ă̓V�[�g���A�N�e�B�u������Ȃ����Ƃ����邪�A
' =           �Ȃ������O��MsgBox����ΑΏ��ł���B
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �V�[�g�I���E�B���h�E��\��()
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Dim bMsgBoxShow As Boolean
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bSHTSELWIN_MSGBOX_SHOW", bMsgBoxShow, bSHTSELWIN_MSGBOX_SHOW, True)
    If bMsgBoxShow = True Then
        MsgBox "�V�[�g�I���E�B���h�E��\�����܂�", vbOKOnly, "�V�[�g�I���E�B���h�E�\��"
    Else
        'Do Nothing
    End If
    
    Application.ScreenUpdating = False
    With CommandBars.Add(Temporary:=True)
        .Controls.Add(ID:=957).Execute
        .Delete
    End With
    Application.ScreenUpdating = True
End Sub

' ==================================================================
' = �T�v    �I��͈͂�͈͂��ێ������܂܃Z���R�s�[����B(�_�u���N�I�[�e�[�V����������)
' = �o��    �E�Z�����ɉ��s���܂܂��ꍇ�͔͈͂�����邱�Ƃɒ���
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' =         SettingFile.cls
' = ����    Macros.bas
' ==================================================================
Public Sub �͈͂��ێ������܂܃Z���R�s�[()
    Const sMACRO_NAME As String = "�͈͂��ێ������܂܃Z���R�s�["
    
    Application.ScreenUpdating = False
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim bIgnoreInvisible As Boolean
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCELLCOPYRNG_IGNORE_INVISIBLE_CELL", bIgnoreInvisible, bCELLCOPYRNG_IGNORE_INVISIBLE_CELL, True)
    
    Dim sDelimiter As String
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYRNG_DELIMITER", sDelimiter, sCELLCOPYRNG_DELIMITER, True)
    
    '*** �I��͈͎擾 ***
    Dim sClipedText As String
    sClipedText = ""
    Dim lAreaIdx As Long
    For lAreaIdx = 1 To Selection.Areas.Count
        '*** �ǉ��e�L�X�g�擾 ***
        Dim asLine() As String
        Call ConvRange2Array( _
            Selection.Areas(lAreaIdx), _
            asLine, _
            bIgnoreInvisible, _
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
        
        If lAreaIdx = 1 Then
            sClipedText = sNewText
        Else
            sClipedText = sClipedText & vbNewLine & sNewText
        End If
    Next lAreaIdx
    
    '*** �N���b�v�{�[�h�ݒ� ***
    Call SetToClipboard(sClipedText)
    
    Application.ScreenUpdating = True
    
    '*** �t�B�[�h�o�b�N ***
    Application.StatusBar = "���������������� " & sMACRO_NAME & "�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v    �I��͈͂���s�ɂ܂Ƃ߂ăZ���R�s�[����B
' = �o��    �E�Z�����ɉ��s���܂܂��ꍇ�͈�s�ɂ܂Ƃ߂��Ȃ����Ƃɒ���
' = �ˑ�    Mng_Clipboard.bas/SetToClipboard()
' =         SettingFile.cls
' = ����    Macro.bas
' =============================================================================
Public Sub ��s�ɂ܂Ƃ߂ăZ���R�s�[()
    Const sMACRO_NAME As String = "��s�ɂ܂Ƃ߂ăZ���R�s�["
    
    Application.ScreenUpdating = False
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim bIgnoreInvisibleCell As Boolean
    Dim bIgnoreBlankCell As Boolean
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCELLCOPYLINE_IGNORE_INVISIBLE_CELL", bIgnoreInvisibleCell, bCELLCOPYLINE_IGNORE_INVISIBLE_CELL, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCELLCOPYLINE_IGNORE_BLANK_CELL", bIgnoreBlankCell, bCELLCOPYLINE_IGNORE_BLANK_CELL, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_PREFFIX", sPreffix, sCELLCOPYLINE_PREFFIX, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_DELIMITER", sDelimiter, sCELLCOPYLINE_DELIMITER, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_SUFFIX", sSuffix, sCELLCOPYLINE_SUFFIX, True)
    
    '*** �I��͈͎擾 ***
    Dim sClipedText As String
    sClipedText = ""
    Dim lAreaIdx As Long
    For lAreaIdx = 1 To Selection.Areas.Count
        Dim lItemIdx As Long
        For lItemIdx = 1 To Selection.Areas(lAreaIdx).Count
            With Selection.Areas(lAreaIdx).Item(lItemIdx)
                If .Value = "" Then                                     '�󔒃Z��
                    If bIgnoreBlankCell = True Then
                        'Do Nothing
                    Else
                        If sClipedText = "" Then
                            sClipedText = sPreffix & .Value
                        Else
                            sClipedText = sClipedText & sDelimiter & .Value
                        End If
                    End If
                Else
                    If .EntireRow.Hidden Or .EntireColumn.Hidden Then   '��\���Z��
                        If bIgnoreInvisibleCell = True Then
                            'Do Nothing
                        Else
                            If sClipedText = "" Then
                                sClipedText = sPreffix & .Value
                            Else
                                sClipedText = sClipedText & sDelimiter & .Value
                            End If
                        End If
                    Else                                                '��L�ȊO
                        If sClipedText = "" Then
                            sClipedText = sPreffix & .Value
                        Else
                            sClipedText = sClipedText & sDelimiter & .Value
                        End If
                    End If
                End If
            End With
        Next lItemIdx
    Next lAreaIdx
    sClipedText = sClipedText & sSuffix
    
    '*** �N���b�v�{�[�h�ݒ� ***
    Call SetToClipboard(sClipedText)
    
    Application.ScreenUpdating = True
    
    '*** �t�B�[�h�o�b�N ***
    Application.StatusBar = "���������������� " & sMACRO_NAME & "�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v    ��s�ɂ܂Ƃ߂ăZ���R�s�[�ɂĎg�p����u�擪����,��؂蕶��,���������v��ύX����
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macro.bas
' =============================================================================
Public Sub ���ݒ�ύX����s�ɂ܂Ƃ߂ăZ���R�s�[()
    Const sMACRO_NAME As String = "���ݒ�ύX����s�ɂ܂Ƃ߂ăZ���R�s�["
    
    Application.ScreenUpdating = False
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_PREFFIX", sPreffix, sCELLCOPYLINE_PREFFIX, False)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_DELIMITER", sDelimiter, sCELLCOPYLINE_DELIMITER, False)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sCELLCOPYLINE_SUFFIX", sSuffix, sCELLCOPYLINE_SUFFIX, False)
    
    Dim vRet As Variant
    vRet = MsgBox( _
        "�u" & sMACRO_NAME & "�v�̐ݒ��ύX���܂��B" & vbNewLine & _
        "�@�擪�����F" & sPreffix & vbNewLine & _
        "�@��؂蕶���F" & sDelimiter & vbNewLine & _
        "�@���������F" & sSuffix & vbNewLine & _
        "" & vbNewLine & _
        "�V���ɐݒ��ύX���܂����H(���͂�)" & vbNewLine & _
        "�f�t�H���g�̐ݒ�ɖ߂��܂����H(��������)", _
        vbYesNoCancel, _
        sMACRO_NAME _
    )
    If vRet = vbYes Then
        sPreffix = InputBox( _
            "�u�擪�����v���w�肵�Ă�������", _
            sMACRO_NAME, _
            sPreffix _
        )
        sDelimiter = InputBox( _
            "�u��؂蕶���v���w�肵�Ă�������", _
            sMACRO_NAME, _
            sDelimiter _
        )
        sSuffix = InputBox( _
            "�u���������v���w�肵�Ă�������", _
            sMACRO_NAME, _
            sSuffix _
        )
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_PREFFIX", sPreffix)
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_SUFFIX", sSuffix)
        MsgBox _
            "�ݒ��ύX���܂���" & vbNewLine & _
            "�@�擪�����F" & sPreffix & vbNewLine & _
            "�@��؂蕶���F" & sDelimiter & vbNewLine & _
            "�@���������F" & sSuffix, _
            vbOKOnly, _
            sMACRO_NAME
    ElseIf vRet = vbNo Then
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_PREFFIX", sPreffix)
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call clSetting.WriteItemToFile(sSettingFilePath, "sCELLCOPYLINE_SUFFIX", sSuffix)
        Application.ScreenUpdating = True
        MsgBox _
            "�ݒ���f�t�H���g�ɖ߂��܂���" & vbNewLine & _
            "�@�擪�����F" & sPreffix & vbNewLine & _
            "�@��؂蕶���F" & sDelimiter & vbNewLine & _
            "�@���������F" & sSuffix, _
            vbOKOnly, _
            sMACRO_NAME
    Else
        Application.ScreenUpdating = True
        MsgBox "�������L�����Z�����܂�", vbExclamation, sMACRO_NAME
    End If
End Sub

' =============================================================================
' = �T�v    �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =         �ׂ荇������̃Z���ɂ̓^�u������}�����ďo�͂���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_Array.bas/ConvRange2Array()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub �t�@�C���G�N�X�|�[�g()
    Const sMACRO_NAME As String = "�t�@�C���G�N�X�|�[�g"
    
    Dim dicDelimiter As Object
    Set dicDelimiter = CreateObject("Scripting.Dictionary")
    
    '�������ݒ聥����
    dicDelimiter.Add "csv", ","
    dicDelimiter.Add "tsv", vbTab
    '�������ݒ聣����
    
    '*** �Z���I�𔻒� ***
    If Selection.Count = 0 Then
        MsgBox "�Z�����I������Ă��܂���", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** �A�h�C���ݒ�t�@�C���p�X�擾 ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim bIgnoreInvisibleCell As Boolean
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bFILEEXPORT_IGNORE_INVISIBLE_CELL", bIgnoreInvisibleCell, bFILEEXPORT_IGNORE_INVISIBLE_CELL, True)
    
    '*** �o�͐���� ***
    '�t�H���_�p�X
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputDirPathInit As String
    Dim sOutputDirPath As String
    sOutputDirPathInit = objWshShell.SpecialFolders("Desktop")
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sFILEEXPORT_OUT_DIR_PATH", sOutputDirPath, sOutputDirPathInit, False)
    sOutputDirPath = ShowFolderSelectDialog(sOutputDirPath)
    If sOutputDirPath = "" Then
        MsgBox "�����ȃt�H���_���w��������̓t�H���_���I������܂���ł����B", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        End
    Else
        'Do Nothing
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sFILEEXPORT_OUT_DIR_PATH", sOutputDirPath)
    
    '�t�@�C����
    Dim sOutputFileName As String
    Dim sOutputFilePath As String
    Dim sFileExt As String
    Dim sDelimiter As String
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sFILEEXPORT_OUT_FILE_NAME", sOutputFileName, sFILEEXPORT_OUT_FILE_NAME, False)
    sOutputFileName = InputBox("�t�@�C��������͂��Ă��������B(�g���q�t��)", sMACRO_NAME, sOutputFileName)
    If InStr(sOutputFileName, ".") Then
        'Do Nothing
    Else
        MsgBox "�t�@�C�������w�肳��܂���ł����B", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        End
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sFILEEXPORT_OUT_FILE_NAME", sOutputFileName)
    
    '�t�@�C���p�X
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName
    
    '*** �g���q,�f���~�^�擾 ***
    sFileExt = Split(sOutputFileName, ".")(UBound(Split(sOutputFileName, ".")))
    If dicDelimiter.Exists(sFileExt) Then
        sDelimiter = dicDelimiter.Item(sFileExt)
    Else
        sDelimiter = vbTab
    End If
    
    '*** �t�@�C���㏑������ ***
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sOutputFilePath) Then
        Dim vAnswer As Variant
        vAnswer = MsgBox("�t�@�C�������݂��܂��B�㏑�����܂����H", vbOKCancel, sMACRO_NAME)
        If vAnswer = vbOK Then
            'Do Nothing
        Else
            MsgBox "�����𒆒f���܂��B", vbExclamation, sMACRO_NAME
            End
        End If
    Else
        'Do Nothing
    End If
    
    '*** �t�@�C���o�͏��� ***
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                sDelimiter _
            )
    
    On Error Resume Next
    Open sOutputFilePath For Output As #1
    If Err.Number = 0 Then
        'Do Nothing
    Else
        MsgBox "�����ȃt�@�C���p�X���w�肳��܂���" & Err.Description, vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        End
    End If
    On Error GoTo 0
    Dim lLineIdx As Long
    For lLineIdx = LBound(asRange) To UBound(asRange)
        Print #1, asRange(lLineIdx)
    Next lLineIdx
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
' = �T�v    �I��͈͓���DOS�R�}���h���o�b�`�t�@�C���ɏ����o���Ă܂Ƃ߂Ď��s����B
' =         �P���I�����̂ݗL���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub DOS�R�}���h���ꊇ���s()
    Const sMACRO_NAME As String = "DOS�R�}���h���ꊇ���s"
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim bIgnoreInvisibleCell As Boolean
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCMDEXEBAT_IGNORE_INVISIBLE_CELL", bIgnoreInvisibleCell, bCMDEXEBAT_IGNORE_INVISIBLE_CELL, True)
    
    '*** �Z���I�𔻒� ***
    If Selection.Count = 0 Then
        MsgBox "�Z�����I������Ă��܂���", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** �͈̓`�F�b�N ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "�P���̂ݑI�����Ă�������", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        End
    End If
    
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                "" _
            )
    
    Dim sBatFileDirPath As String
    Dim sBatFilePath As String
    sBatFileDirPath = GetAddinSettingDirPath()
    sBatFilePath = sBatFileDirPath & "\" & sCMDEXEBAT_BAT_FILE_NAME
    Debug.Print sBatFilePath
    
    Call OutputTxtFile(sBatFilePath, asRange)
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEBAT_REDIRECT_FILE_NAME
    
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
    
    MsgBox "���s�����I", vbOKOnly, sMACRO_NAME
    
    '*** �o�̓t�@�C�����J�� ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v    �I��͈͓���DOS�R�}���h�����ꂼ����s����B
' =         �P���I�����̂ݗL���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub DOS�R�}���h���e�X���s()
    Const sMACRO_NAME As String = "DOS�R�}���h���e�X���s"
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim bIgnoreInvisibleCell As Boolean
    Call clSetting.ReadItemFromFile(sSettingFilePath, "bCMDEXEUNI_IGNORE_INVISIBLE_CELL", bIgnoreInvisibleCell, bCMDEXEUNI_IGNORE_INVISIBLE_CELL, True)
    
    '*** �Z���I�𔻒� ***
    If Selection.Count = 0 Then
        MsgBox "�Z�����I������Ă��܂���", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** �͈̓`�F�b�N ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "�P���̂ݑI�����Ă�������", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        End
    End If
    
    'Range�^����String()�^�֕ϊ�
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                "" _
            )
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEUNI_REDIRECT_FILE_NAME
    
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
    
    MsgBox "���s�����I", vbOKOnly, sMACRO_NAME
    
    '*** �o�̓t�@�C�����J�� ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v    �I��͈͓��̌��������̕����F��ύX����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub ���������̕����F��ύX()
    Const sMACRO_NAME As String = "���������̕����F��ύX"
    Const lSELECT_CLR_PALETTE As Boolean = True
    
    '�������F�ݒ聥����
    Const sCOLOR_TYPE As String = "0:�ԁA1:���A2:�΁A3:���A4:��A5:���A6:���A7:��"
    Dim cCLR_RGBS As Variant
    Set cCLR_RGBS = CreateObject("System.Collections.ArrayList")
    cCLR_RGBS.Add &HFF
    cCLR_RGBS.Add &HC6AC4B
    cCLR_RGBS.Add &H3C9376
    cCLR_RGBS.Add &HA03070
    cCLR_RGBS.Add &H4696F7
    cCLR_RGBS.Add &HC0FF
    cCLR_RGBS.Add &HFFFFFF
    cCLR_RGBS.Add &H0
    '�������F�ݒ聣����
    
    '*** �A�h�C���ݒ�t�@�C������ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Dim sSrchStr As String
    Dim lClrRgbInit As Long
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sWORDCOLOR_SRCH_WORD", sSrchStr, sWORDCOLOR_SRCH_WORD, False)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lWORDCOLOR_CLR_RGB", lClrRgbInit, lWORDCOLOR_CLR_RGB, False)
    
    '�����Ώە�����I��
    sSrchStr = InputBox("�������������͂��Ă�������", sMACRO_NAME, sSrchStr)
    If StrPtr(sSrchStr) = 0 Then
        MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    ElseIf sSrchStr = "" Then
        MsgBox "�����񂪎w�肳��Ȃ��������߁A�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    
    '�F�I��
    Dim lClrRgbSelected As Long
    If lSELECT_CLR_PALETTE = True Then '�J���[�p���b�g�őI��
        Dim bRet As Boolean
        bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
        If bRet = False Then
            MsgBox "�F�I�������s���܂����̂ŁA�����𒆒f���܂��B", vbCritical, sMACRO_NAME
            Exit Sub
        End If
    Else '�F��ʖ��őI��
        '�F���F��� �ϊ�
        Dim lClrTypeIdx As Long
        lClrTypeIdx = 0
        Dim bExist As Boolean
        bExist = False
        Dim vClrRgb As Variant
        For Each vClrRgb In cCLR_RGBS
            If vClrRgb = lClrRgbInit Then
                bExist = True
                Exit For
            Else
                lClrTypeIdx = lClrTypeIdx + 1
            End If
        Next
        If bExist = True Then
            'Do Nothing
        Else
            lClrTypeIdx = 0
        End If
        '�F��� �I��
        lClrTypeIdx = InputBox( _
            "�����F��I�����Ă�������" & vbNewLine & _
            "  " & sCOLOR_TYPE & vbNewLine _
            , _
            sMACRO_NAME, _
            lClrTypeIdx _
        )
        '�F��ʁ��F �ϊ�
        If lClrTypeIdx < cCLR_RGBS.Count Then
            lClrRgbSelected = cCLR_RGBS(lClrTypeIdx)
        Else
            MsgBox "�����F�͎w��͈͓̔��őI�����Ă��������B" & vbNewLine & sCOLOR_TYPE, vbOKOnly, sMACRO_NAME
            Exit Sub
        End If
    End If
    
    '�A�h�C���ݒ�X�V
    Call clSetting.WriteItemToFile(sSettingFilePath, "sWORDCOLOR_SRCH_WORD", sSrchStr)
    Call clSetting.WriteItemToFile(sSettingFilePath, "lWORDCOLOR_CLR_RGB", lClrRgbSelected)
    
    '�Ώ۔͈͓���(�I��͈͂Ǝg�p����Ă���͈͂̋��ʕ���)
    Dim rTrgtRng As Range
    Set rTrgtRng = Application.Intersect(Selection, ActiveSheet.UsedRange)
    
    '����������F�ύX
    Dim oCell As Range
    For Each oCell In rTrgtRng
        Dim sTrgtStr As String
        sTrgtStr = oCell.Value
        Dim lStartIdx As Long
        lStartIdx = 1
        Do While True '�������[�v
            Dim lIdx As Long
            lIdx = InStr(lStartIdx, sTrgtStr, sSrchStr)
            If lIdx = 0 Then
                Exit Do
            Else
                lStartIdx = lIdx + Len(sSrchStr)
                oCell.Characters(Start:=lIdx, Length:=Len(sSrchStr)).Font.Color = lClrRgbSelected
            End If
        Loop
    Next
    
    MsgBox "�����I", vbOKOnly, sMACRO_NAME
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
    Const sMACRO_NAME As String = "�V�[�g���בւ���Ɨp�V�[�g���쐬"
    Const WORK_SHEET_NAME As String = "�V�[�g���בւ���Ɨp"
    Const ROW_BTN = 2
    Const ROW_TEXT_1 = 4
    Const ROW_TEXT_2 = 5
    Const ROW_SHT_NAME_TITLE = 7
    Const ROW_SHT_NAME_STRT = 8
    Const CLM_BTN = 2
    Const CLM_SHT_NAME = 2
    
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
            MsgBox "���Ɂu" & WORK_SHEET_NAME & "�v�V�[�g���쐬����Ă��܂��B", vbCritical, sMACRO_NAME
            MsgBox "�����𑱂������ꍇ�́A�V�[�g���폜���Ă��������B", vbCritical, sMACRO_NAME
            MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
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

' ==================================================================
' = �T�v    �I���V�[�g��ʃt�@�C���ɐ؂�o���B
' =         �R�s�[���u�b�N�Ɠ��t�H���_�ɏo�͂���B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Public Sub �I���V�[�g�؂�o��()
    Const sMACRO_NAME As String = "�I���V�[�g�؂�o��"
    
    Dim shSht As Worksheet
    Dim wSrcWindow As Window
    Dim bkSrcBook As Workbook
    Dim bkTrgtBook As Workbook
    Dim sTrgtBookName As String
    
    Set bkSrcBook = ActiveWorkbook
    Set wSrcWindow = ActiveWindow
    Set bkTrgtBook = Workbooks.Add
    
    wSrcWindow.SelectedSheets.Copy _
        After:=bkTrgtBook.Sheets(bkTrgtBook.Sheets.Count)
    Application.DisplayAlerts = False
    bkTrgtBook.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    bkTrgtBook.SaveAs bkSrcBook.Path & "\" & wSrcWindow.SelectedSheets(1).Name & ".xlsx"
    bkTrgtBook.Close
    
    MsgBox "�I���V�[�g�؂�o�������I", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = �T�v    �I�������͈͂̃n�C�p�[�����N���ꊇ�ŊJ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �n�C�p�[�����N�ꊇ�I�[�v��()
    Const sMACRO_NAME As String = "�n�C�p�[�����N�ꊇ�I�[�v��"
    Dim Rng As Range
    
    If TypeName(Selection) = "Range" Then
        For Each Rng In Selection
            If Rng.Hyperlinks.Count > 0 Then Rng.Hyperlinks(1).Follow
        Next
    Else
        MsgBox "�Z���͈͂��I������Ă��܂���B", vbExclamation, sMACRO_NAME
    End If
End Sub

' =============================================================================
' = �T�v    �t�H���g�F���u�ݒ�F�v�́u�����v�Ńg�O������
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub �t�H���g�F���g�O��()
    '�A�h�C���ݒ�ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim lClrRgb As Long
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lCLRTGLFONT_CLR_RGB", lClrRgb, lCLRTGLFONT_CLR_RGB, True)
    
    '�t�H���g�F�ύX
    If Selection(1).Font.Color = lClrRgb Then
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Font.Color = lClrRgb
    End If
End Sub

' =============================================================================
' = �T�v    �w�i�F���u�ݒ�F�v�́u�w�i�F�Ȃ��v�Ńg�O������
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub �w�i�F���g�O��()
    '�A�h�C���ݒ�ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim lClrRgb As Long
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lCLRTGLBG_CLR_RGB", lClrRgb, lCLRTGLBG_CLR_RGB, True)
    
    '�w�i�F�ύX
    If Selection(1).Interior.Color = lClrRgb Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = lClrRgb
    End If
End Sub

' =============================================================================
' = �T�v    �u�t�H���g�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' =         Macros.bas/ShowColorPalette()
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���t�H���g�F���g�O���̐F�I��()
    Const sMACRO_NAME As String = "���ݒ�ύX���t�H���g�F���g�O���̐F�I��"
    
    MsgBox sMACRO_NAME & "�����s���܂�", vbOKOnly, sMACRO_NAME
    
    '�A�h�C���ݒ�ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim lClrRgbInit As Long
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lCLRTGLFONT_CLR_RGB", lClrRgbInit, lCLRTGLFONT_CLR_RGB, False)
    
    '�F�I��
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "�F�I�������s���܂����̂ŁA�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    '�A�h�C���ݒ�X�V
    Call clSetting.WriteItemToFile(sSettingFilePath, "lCLRTGLFONT_CLR_RGB", lClrRgbSelected)
End Sub

' =============================================================================
' = �T�v    �u�w�i�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' =         Macros.bas/ShowColorPalette()
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���w�i�F���g�O���̐F�I��()
    Const sMACRO_NAME As String = "���ݒ�ύX���w�i�F���g�O���̐F�I��"
    
    MsgBox sMACRO_NAME & "�����s���܂�", vbOKOnly, sMACRO_NAME
    
    '�A�h�C���ݒ�ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim lClrRgbInit As Long
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lCLRTGLBG_CLR_RGB", lClrRgbInit, lCLRTGLBG_CLR_RGB, False)
    
    '�F�I��
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "�F�I�������s���܂����̂ŁA�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    '�A�h�C���ݒ�X�V
    Call clSetting.WriteItemToFile(sSettingFilePath, "lCLRTGLBG_CLR_RGB", lClrRgbSelected)
End Sub

' =============================================================================
' = �T�v    �u�t�H���g�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g()
    Const sMACRO_NAME As String = "���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g"
    
    '�F�擾
    Dim lClrRgb As Long
    lClrRgb = Selection(1).Font.Color
    
    '�A�h�C���ݒ�X�V
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.WriteItemToFile(sSettingFilePath, "lCLRTGLFONT_CLR_RGB", lClrRgb)
End Sub

' =============================================================================
' = �T�v    �u�w�i�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g()
    Const sMACRO_NAME As String = "���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g"
    
    '�F�擾
    Dim lClrRgb As Long
    lClrRgb = Selection(1).Interior.Color
    
    '�A�h�C���ݒ�X�V
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.WriteItemToFile(sSettingFilePath, "lCLRTGLBG_CLR_RGB", lClrRgb)
End Sub

' =============================================================================
' = �T�v    �C���f���g���グ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �C���f���g���グ��()
    Dim rCell As Range
    For Each rCell In Selection
        rCell.IndentLevel = rCell.IndentLevel + 1
    Next
End Sub

' =============================================================================
' = �T�v    �C���f���g��������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �C���f���g��������()
    Dim rCell As Range
    For Each rCell In Selection
        If rCell.IndentLevel = 0 Then
            'Do Nothing
        Else
            rCell.IndentLevel = rCell.IndentLevel - 1
        End If
    Next
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
' = �T�v    �A�N�e�B�u�u�b�N�̐擪�V�[�g�ֈړ�����
' = �o��    �Ȃ�
' = �ˑ��@�@�Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �擪�V�[�g�փW�����v()
    Application.ScreenUpdating = False
    Dim shSheet As Worksheet
    For Each shSheet In ActiveWorkbook.Sheets
        If shSheet.Visible = True Then
            shSheet.Activate
            Exit For
        End If
    Next
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N�̖����V�[�g�ֈړ�����
' = �o��    �Ȃ�
' = �ˑ��@�@�Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �����V�[�g�փW�����v()
    Application.ScreenUpdating = False
    With ActiveWorkbook
        Dim lShtCnt As Long
        For lShtCnt = .Sheets.Count To 1 Step -1
            If .Sheets(lShtCnt).Visible = True Then
                .Sheets(lShtCnt).Activate
                Exit For
            End If
        Next
    End With
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�Z���̃R�����g�\���̗L��/������؂�ւ���
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' =         Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��()
    Const sMACRO_NAME As String = "���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��"
    
    '�A�h�C���ݒ�t�@�C���ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Dim bCmntVsblEnb As Boolean
    Dim bExistSetting As Boolean
    bExistSetting = clSetting.ReadItemFromFile(sSettingFilePath, "bCMNT_VSBL_ENB", bCmntVsblEnb, bCMNT_VSBL_ENB, False)
    
    '�A�N�e�B�u�Z���R�����g�ݒ�X�V
    If bExistSetting = True Then
        If bCmntVsblEnb = True Then
            MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\�����y�������z���܂�", vbOKOnly, sMACRO_NAME
            bCmntVsblEnb = False
        Else
            MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\�����y�L�����z���܂�", vbOKOnly, sMACRO_NAME
            bCmntVsblEnb = True
        End If
    Else
        MsgBox "�A�N�e�B�u�Z���R�����g�̂ݕ\�����y�L�����z���܂�", vbOKOnly, sMACRO_NAME
        bCmntVsblEnb = True
    End If
    
    Call clSetting.WriteItemToFile(sSettingFilePath, "bCMNT_VSBL_ENB", bCmntVsblEnb)
    
    '�V���[�g�J�b�g�L�[�ݒ� �X�V(�L����)
    Call SwitchMacroShortcutKeysActivation(True)
End Sub

' =============================================================================
' = �T�v    ���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h(+�ړ�)
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/VisibleCommentOnlyActiveCell()
' = ����    Macros.bas
' =============================================================================
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\��()
'   Application.ScreenUpdating = False
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉ��ړ�()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(1, 0).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\�����ď�ړ�()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(-1, 0).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\�����č��ړ�()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(0, -1).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉE�ړ�()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(0, 1).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    Excel���ᎆ
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub Excel���ᎆ()
    '�A�h�C���ݒ�ǂݏo��
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim sFontName As String
    Dim lFontSize As Long
    Dim lClmWidth As Long
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEXCELGRID_FONT_NAME", sFontName, sEXCELGRID_FONT_NAME, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lEXCELGRID_FONT_SIZE", lFontSize, lEXCELGRID_FONT_SIZE, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lEXCELGRID_CLM_WIDTH", lClmWidth, lEXCELGRID_CLM_WIDTH, True)
    
    'Excel���ᎆ�ݒ�
    ActiveSheet.Cells.Select
    With Selection
        .Font.Name = sFontName
        .Font.Size = lFontSize
        .ColumnWidth = lClmWidth
        .Rows.AutoFit
    End With
    ActiveSheet.Cells(1, 1).Select
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
' = �ˑ�    Mng_FileSys.bas/ShowFileSelectDialog()
' =         Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_Collection.bas/ReadTxtFileToCollection()
' =         Mng_String.bas/ExecRegExp()
' =         Mng_String.bas/ExtractTailWord()
' =         Mng_String.bas/ExtractRelativePath()
' =         Mng_ExcelOpe.bas/CreateNewWorksheet()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub EpTree�̊֐��c���[��Excel�Ŏ�荞��()
    Const sMACRO_NAME As String = "EpTree�̊֐��c���[��Excel�Ŏ�荞��"
    Const lSTRT_ROW As Long = 1
    Const lSTRT_CLM As Long = 1
    
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    '=============================================
    '= ���O����
    '=============================================
    Dim sOutSheetName As String
    Dim lMaxFuncLevelIni As Long
    Dim lClmWidth As Long
    Dim sEptreeLogPath As String
    Dim sDevRootDirPath As String
    Dim sDevRootDirName As String
    Dim lDevRootLevel As Long
    
    '*** �A�h�C���ݒ�t�@�C������ݒ�ǂݏo�� ***
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEPTREE_OUT_SHEET_NAME", sOutSheetName, sEPTREE_OUT_SHEET_NAME, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lEPTREE_MAX_FUNC_LEVEL_INI", lMaxFuncLevelIni, lEPTREE_MAX_FUNC_LEVEL_INI, True)
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lEPTREE_CLM_WIDTH", lClmWidth, lEPTREE_CLM_WIDTH, True)
    
    'Eptree���O�t�@�C���p�X�擾
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEPTREE_OUT_LOG_PATH", sEptreeLogPath, sEPTREE_OUT_LOG_PATH, False)
    sEptreeLogPath = ShowFileSelectDialog(sEptreeLogPath, "EpTreeLog.txt�̃t�@�C���p�X��I�����Ă�������")
    If sEptreeLogPath = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "sEPTREE_OUT_LOG_PATH", sEptreeLogPath)
    
    '�J���p���[�g�t�H���_�擾
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEPTREE_DEV_ROOT_DIR_PATH", sDevRootDirPath, sEPTREE_DEV_ROOT_DIR_PATH, False)
    sDevRootDirPath = ShowFolderSelectDialog(sDevRootDirPath, "�J���p���[�g�t�H���_�p�X��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j")
    If sDevRootDirPath = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    sDevRootDirName = ExtractTailWord(sDevRootDirPath, "\")
    Call clSetting.WriteItemToFile(sSettingFilePath, "sEPTREE_DEV_ROOT_DIR_PATH", sDevRootDirPath)
    
    '���[�g�t�H���_���x���擾
    Call clSetting.ReadItemFromFile(sSettingFilePath, "lEPTREE_DEV_ROOT_DIR_LEVEL", lDevRootLevel, lEPTREE_DEV_ROOT_DIR_LEVEL, False)
    Dim sDevRootLevel As String
    sDevRootLevel = InputBox("���[�g�t�H���_���x������͂��Ă�������", sMACRO_NAME, CStr(lDevRootLevel))
    If sDevRootLevel = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call clSetting.WriteItemToFile(sSettingFilePath, "lEPTREE_DEV_ROOT_DIR_LEVEL", sDevRootLevel)
    
    '=============================================
    '= �{����
    '=============================================
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    '�V�[�g�ǉ�
    Dim sSheetName As String
    Dim shTrgtSht As Worksheet
    sSheetName = CreateNewWorksheet(sOutSheetName)
    Set shTrgtSht = ActiveWorkbook.Sheets(sSheetName)
    
    '�e�L�X�g�t�@�C���ǂݏo��
    Dim cFileContents As Collection
    Set cFileContents = New Collection
    Call ReadTxtFileToCollection(sEptreeLogPath, cFileContents)
    
    '�t�@�C���c���[�o��
    lStrtRow = lSTRT_ROW
    lStrtClm = lSTRT_CLM
    lRowIdx = lStrtRow
    
    With shTrgtSht
        .Cells(lRowIdx, lStrtClm + 0).Value = "�t�@�C���p�X"
        .Cells(lRowIdx, lStrtClm + 1).Value = "�s��"
        .Cells(lRowIdx, lStrtClm + 2).Value = "�֐���"
        .Cells(lRowIdx, lStrtClm + 3).Value = "�R�[���c���["
    End With
    lRowIdx = lRowIdx + 1
    
    Dim lMaxFuncLevel As Long
    lMaxFuncLevel = lMaxFuncLevelIni
    Dim vFileLine As Variant
    For Each vFileLine In cFileContents
        Dim oMatchResult As Object
        Call ExecRegExp( _
            vFileLine, _
            "^([^ ]+)? +(\d+): (  )?([��|��|��|  ]*)(\w+)(��)?", _
            oMatchResult _
        )
        
        Dim sFilePath As String
        Dim sLineNo As String
        Dim lFuncLevel As Long
        Dim sFuncName As String
        Dim sOmission As String
        sFilePath = oMatchResult(0).SubMatches(0)
        Call ExtractRelativePath(sFilePath, sDevRootDirName, Int(sDevRootLevel), sFilePath)
        sLineNo = oMatchResult(0).SubMatches(1)
        If sLineNo = 0 Then
            sLineNo = ""
        End If
        lFuncLevel = LenB(StrConv(oMatchResult(0).SubMatches(3), vbFromUnicode)) / 2
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
        lLastClm = lSTRT_CLM + 3 + lMaxFuncLevel
        lLastRow = lRowIdx
        
        '�^�C�g���s ��������
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lStrtRow, lStrtClm + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lStrtRow, lLastClm)).HorizontalAlignment = xlCenterAcrossSelection
        
        '�񕝒���
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lLastRow, lStrtClm + 0)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 1), .Cells(lLastRow, lStrtClm + 1)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 2), .Cells(lLastRow, lStrtClm + 2)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lLastRow, lLastClm)).ColumnWidth = lClmWidth
        
        '�I�[�g�t�B���^
        .Range(.Cells(lStrtRow, lStrtClm), .Cells(lLastRow, lLastClm)).AutoFilter
        
        '�s����
        .Rows(lStrtRow).RowHeight = .Rows(lStrtRow).RowHeight * 3
        
        '�^�C�g����Œ�
        ActiveWindow.FreezePanes = False
        .Rows(lStrtRow + 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
        
        '�V�[�g���o���F
        .Tab.Color = RGB(242, 220, 219)
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "�֐��R�[���c���[�쐬�����I", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = �T�v    Excel�������`�����{/����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub Excel�������`�����{()
    Dim rSelectRange As Range
    For Each rSelectRange In Selection
        rSelectRange.Formula = ConvFormuraIndentation(rSelectRange.Formula, True)
    Next
End Sub
Public Sub Excel�������`������()
    Dim rSelectRange As Range
    For Each rSelectRange In Selection
        rSelectRange.Formula = ConvFormuraIndentation(rSelectRange.Formula, False)
    Next
End Sub

' =============================================================================
' = �T�v    �A�h�C���}�N�����s
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �A�h�C���}�N�����s()
    ExecAddInMacro.Show
End Sub

' *****************************************************************************
' * �����v���V�[�W����`
' *****************************************************************************
Private Sub ���������������v���V�[�W������������()
    '�v���V�[�W�����X�g�\���p�̃_�~�[�v���V�[�W��
End Sub

' =============================================================================
' = �T�v    �V�[�g����ёւ���B
' =         �V�[�g���בւ���Ɨp�V�[�g�ɋL�ڂ̒ʂ�A�V�[�g����ёւ���B
' =         �K���V�[�g���בւ���Ɨp�V�[�g����Ăяo�����ƁI
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Private Sub SortSheetPost()
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
' = �T�v    ���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���B
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Private Sub VisibleCommentOnlyActiveCell()
    On Error Resume Next
    
    '�S�Z���R�����g��\��
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    '�A�N�e�B�u�Z���R�����g�\��
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
End Sub

' ==================================================================
' = �T�v    �A�h�C���ݒ�p�̃t�@�C���p�X���擾����
' = ����    �Ȃ�
' = �ߒl                    String      �A�h�C���ݒ�p�̃t�@�C���p�X
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Public Function GetAddinSettingFilePath() As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    GetAddinSettingFilePath = GetAddinSettingDirPath() & "\" & objFSO.GetBaseName(ThisWorkbook.Name) & ".cfg"
End Function

' ==================================================================
' = �T�v    �A�h�C���ݒ�p�̃t�H���_�p�X���擾����
' = ����    �Ȃ�
' = �ߒl                    String      �A�h�C���ݒ�p�̃t�H���_�p�X
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Public Function GetAddinSettingDirPath() As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    GetAddinSettingDirPath = _
        objWshShell.SpecialFolders("MyDocuments") & "\" & objFSO.GetBaseName(ThisWorkbook.Name)
End Function

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
' = ����    bIgnoreInvisibleCell    String  [in]  ��\���Z���������s��
' = ����    sDelimiter              String  [in]  ��؂蕶��
' = �ߒl    �Ȃ�
' = �o��    �񂪗ׂ荇�����Z�����m�͎w�肳�ꂽ��؂蕶���ŋ�؂���
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Private Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIgnoreInvisibleCell As Boolean, _
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
            If bIgnoreInvisibleCell = True Then
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
' = �T�v    �t�H���_�I���_�C�A���O��\������
' = ����    sInitPath       String  [in]  �f�t�H���g�t�H���_�p�X�i�ȗ��j
' = ����    sTitle          String  [in]  �^�C�g�����i�ȗ��j
' = �ߒl                    String        �I���t�H���_
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If sTitle = "" Then
        fdDialog.Title = "�t�H���_��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j"
    Else
        fdDialog.Title = sTitle
    End If
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

' ==================================================================
' = �T�v    �t�@�C���i�P��j�I���_�C�A���O��\������
' = ����    sInitPath       String  [in]  �f�t�H���g�t�@�C���p�X�i�ȗ��j
' = ����    sTitle          String  [in]  �^�C�g�����i�ȗ��j
' = ����    sFilters        String  [in]  �I�����̃t�B���^�i�ȗ��j(��)
' = �ߒl                    String        �I���t�@�C��
' = �o��    (��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =              ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =                    �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =                    �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =                    �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' = �ˑ�    Mng_FileSys.bas/SetDialogFilters()
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    If sTitle = "" Then
        fdDialog.Title = "�t�@�C����I�����Ă�������"
    Else
        fdDialog.Title = sTitle
    End If
    fdDialog.AllowMultiSelect = False
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
        ShowFileSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
            ShowFileSelectDialog = sSelectedPath
        Else
            ShowFileSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
 
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
        .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).Name = sSheetName
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
' = �T�v    ��΃p�X���猟���L�[�z���K�w�̑��΃p�X�֒u��
' = ����    sInFilePath     String  [in]    ��΃p�X
' = ����    sMatchDirName   String  [in]    �����Ώۃt�H���_��
' = ����    lRemeveDirLevel Long    [in]    �K�w���x��
' = ����    sRelativePath   String  [out]   ���΃p�X
' = �ߒl                    Boolean         ��������
' = �o��    ���s��1)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : codes
'             lRemeveDirLevel : 1
'             ��
'             sRelativePath   : bbb\ccc\test.txt
'             �ߒl            : true
'
'           ���s��2)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : code
'             lRemeveDirLevel : 2
'             ��
'             sRelativePath   : c\codes\aaa\bbb\ccc\test.txt
'             �ߒl            : false
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function ExtractRelativePath( _
    ByVal sInFilePath As String, _
    ByVal sMatchDirName As String, _
    ByVal lRemeveDirLevel As Long, _
    ByRef sRelativePath As String _
) As Boolean
    Dim sRemoveDirLevelPath
    sRemoveDirLevelPath = ""
    Dim lIdx
    For lIdx = 0 To lRemeveDirLevel - 1
        sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
    Next
    
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
    sTargetStr = sInFilePath
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = True                           '������S�̂�����
    
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
    
    If oMatchResult.Count > 0 Then
        sRelativePath = Replace(sInFilePath, oMatchResult.Item(0), "")
        ExtractRelativePath = True
    Else
        sRelativePath = sInFilePath
        ExtractRelativePath = False
    End If
End Function

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim asSplitWord() As String
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function

' ==================================================================
' = �T�v    Excel�����𐮌`����
' = ����    sInputCellFormula   String   [in]   ���͐���
' = ����    bExecIndentation    Boolean  [in]   ���`���{/���`����
' = ����    lIndentWidth        Long     [in]   �C���f���g������(�ȗ���)
' = �ߒl                        String          �o�͐���
' = �o��    �E���`�������́A�����Ɋ֌W�̂Ȃ��󔒂͂��ׂď�������
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Private Function ConvFormuraIndentation( _
    ByVal sInputCellFormula As String, _
    ByVal bExecIndentation As Boolean, _
    Optional ByVal lIndentWidth As Long = 4 _
) As String
    Dim sOutputCellFormula As String
    sOutputCellFormula = ""
    
    '�����̏ꍇ
    If Left(sInputCellFormula, 1) = "=" Then
        Dim bStrMode As Boolean
        Dim lNestCnt As Long
        bStrMode = False
        lNestCnt = 0
        '�����񑀍�
        Dim lChrIdx As Long
        For lChrIdx = 1 To Len(sInputCellFormula)
            Dim sInputCellFormulaChr As String
            sInputCellFormulaChr = Mid(sInputCellFormula, lChrIdx, 1)
            
            '�����񃂁[�h�̏ꍇ
            If bStrMode = True Then
                Select Case sInputCellFormulaChr
                Case """"
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    bStrMode = False
                Case Else
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                End Select
            '�����񃂁[�h�łȂ��ꍇ
            Else
                Select Case sInputCellFormulaChr
                Case ","
                    If bExecIndentation = True Then
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr & vbLf & String(lNestCnt * lIndentWidth, " ")
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case "("
                    If bExecIndentation = True Then
                        lNestCnt = lNestCnt + 1
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr & vbLf & String(lNestCnt * lIndentWidth, " ")
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case ")"
                    If bExecIndentation = True Then
                        lNestCnt = lNestCnt - 1
                        sOutputCellFormula = sOutputCellFormula & vbLf & String(lNestCnt * lIndentWidth, " ") & sInputCellFormulaChr
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case """"
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    bStrMode = True
                Case vbLf
                    'Do Nothing
                Case " "
                    'Do Nothing
                Case Else
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                End Select
            End If
        Next lChrIdx
    '�����łȂ��ꍇ
    Else
        sOutputCellFormula = sInputCellFormula
    End If
    
    ConvFormuraIndentation = sOutputCellFormula
End Function

' ==================================================================
' = �T�v    �F�̐ݒ�_�C�A���O��\�����A�����őI�����ꂽ�F��RGB�l��Ԃ�
' = ����    lClrRgbInit       Long    [in]    RGB�l �����l
' = ����    lClrRgbSelected   Long    [out]   RGB�l �I��l
' = �ߒl                      Boolean         �I������
' =                                               (True:����,False:�L�����Z��or���s)
' = �o��    �E�L�����Z��or���s���AlClrRgbSelected��Init�Ɠ����l�ƂȂ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Private Function ShowColorPalette( _
    ByVal lClrRgbInit As Long, _
    ByRef lClrRgbSelected As Long _
) As Boolean
    Const CC_RGBINIT = &H1          '�F�̃f�t�H���g�l��ݒ�
    Const CC_LFULLOPEN = &H2        '�F�̍쐬���s��������\��
    Const CC_PREVENTFULLOPEN = &H4  '�F�̍쐬�{�^���𖳌��ɂ���
    Const CC_SHOWHELP = &H8         '�w���v�{�^����\��
    
    Dim tChooseColor As ChooseColor
    With tChooseColor
        '�_�C�A���O�̐ݒ�
        .lStructSize = Len(tChooseColor)
        .lpCustColors = String$(64, Chr$(0))
        .flags = CC_RGBINIT + CC_LFULLOPEN
        .rgbResult = lClrRgbInit
        
        '�_�C�A���O��\��
        Dim lRet As Long
        lRet = ChooseColor(tChooseColor)
        
        '�_�C�A���O����̕Ԃ�l���`�F�b�N
        lClrRgbSelected = lClrRgbInit
        If lRet <> 0 Then
            If .rgbResult > RGB(255, 255, 255) Then '�G���[
                ShowColorPalette = False
            Else '����I��
                ShowColorPalette = True
                lClrRgbSelected = .rgbResult
            End If
        Else '�L�����Z������
            ShowColorPalette = False
        End If
    End With
End Function

