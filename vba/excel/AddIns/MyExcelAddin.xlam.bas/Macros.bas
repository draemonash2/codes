Attribute VB_Name = "Macros"
Option Explicit

' my excel addin macros v2.24

' =============================================================================
' =  <<�}�N���ꗗ>>
' =     �E����
' =         F1�w���v������                              F1�w���v�𖳌�������
' =
' =     �E�}�N���ݒ�
' =         �}�N���V���[�g�J�b�g�L�[�S�ėL����          �}�N���V���[�g�J�b�g�L�[�S�ėL����
' =         �}�N���V���[�g�J�b�g�L�[�S�Ė�����          �}�N���V���[�g�J�b�g�L�[�S�Ė�����
' =         �A�h�C���}�N�����s                          �A�h�C���}�N�����s
' =         ���W���[���ꊇ�G�N�X�|�[�g_�A�h�C��         �{�A�h�C�����̑S�}�N��/�v���V�[�W�����G�N�X�|�[�g����
' =         ���W���[���ꊇ�G�N�X�|�[�g_�A�N�e�B�u�u�b�N �A�N�e�B�u�u�b�N���̑S�}�N��/�v���V�[�W�����G�N�X�|�[�g����
' =         CtrlShiftF�}�N��                            �V���[�g�J�b�g�L�[�d�����̐U�蕪������(Ctrl + Shift + F)
' =
' =     �E�u�b�N����
' =         �ʃv���Z�X�ŊJ��                            �A�N�e�B�u�u�b�N��ʃv���Z�X�ŊJ��
' =         �t�@�C���p�X�R�s�[                          �A�N�e�B�u�u�b�N�̃t�@�C���p�X���R�s�[
' =         �t�@�C�����R�s�[                            �A�N�e�B�u�u�b�N�̃t�@�C�������R�s�[
' =
' =     �E�V�[�g����
' =         EpTree�̊֐��c���[��Excel�Ŏ�荞��         EpTree�̊֐��c���[��Excel�Ŏ�荞��
' =         Excel���ᎆ                                 Excel���ᎆ
' =         �I���V�[�g�؂�o��                          �I���V�[�g��ʃt�@�C���ɐ؂�o��
' =         �S�V�[�g�����R�s�[                          �u�b�N���̃V�[�g����S�ăR�s�[����
' =         �V�[�g�\����\����؂�ւ�                  �V�[�g�\��/��\����؂�ւ���
' =         �V�[�g���בւ���Ɨp�V�[�g���쐬            �V�[�g���בւ���Ɨp�V�[�g�쐬
' =         �V�[�g�I���E�B���h�E��\��                  �V�[�g�I���E�B���h�E��\������
' =         �V�[�g���ꊇ�ύX                            �V�[�g�����ꊇ�ύX����
' =         �V�[�g�ǉ��J�X�^��                          �V�[�g��ǉ�����i�J�X�^���ݒ�Łj
' =         �擪�V�[�g�փW�����v                        �A�N�e�B�u�u�b�N�̐擪�V�[�g�ֈړ�����
' =         �����V�[�g�փW�����v                        �A�N�e�B�u�u�b�N�̖����V�[�g�ֈړ�����
' =         �V�[�g�Čv�Z���Ԍv��                        �V�[�g���ɍČv�Z�ɂ����鎞�Ԃ��v������
' =
' =     �E�Z������
' =         �t�@�C���G�N�X�|�[�g                        �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =         DOS�R�}���h���ꊇ���s                       �I��͈͓���DOS�R�}���h���܂Ƃ߂Ď��s����B
' =         DOS�R�}���h���e�X���s                       �I��͈͓���DOS�R�}���h�����ꂼ����s����B
' =         DOS�R�}���h���ꊇ���s_�Ǘ��Ҍ���            �I��͈͓���DOS�R�}���h���܂Ƃ߂Ď��s����B�i�Ǘ��Ҍ����j
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
' =         �N���b�v�{�[�h�l�\��t��                    �N���b�v�{�[�h����l�\��t������
' =         �t�H���g�F���g�O��                          �t�H���g�F���u�ݒ�F�v�́u�����v�Ńg�O������
' =         ���ݒ�ύX���t�H���g�F���g�O���̐F�I��      �u�t�H���g�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' =         ���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g  �u�t�H���g�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' =         �w�i�F���g�O��                              �w�i�F���u�ݒ�F�v�́u�w�i�F�Ȃ��v�Ńg�O������
' =         ���ݒ�ύX���w�i�F���g�O���̐F�I��          �u�w�i�F���g�O���v�̐ݒ�F���J���[�p���b�g����擾���ĕύX����
' =         ���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g      �u�w�i�F���g�O���v�̐ݒ�F���A�N�e�B�u�Z������擾���ĕύX����
' =         �I�[�g�t�B�����s                            �I�[�g�t�B�������s����
' =         ��ʂ���Ɉړ�                              ��ʂ���Ɉړ�(�X�N���[�����b�N����)
' =         ��ʂ����Ɉړ�                              ��ʂ����Ɉړ�(�X�N���[�����b�N����)
' =         ��ʂ����Ɉړ�                              ��ʂ����Ɉړ�(�X�N���[�����b�N����)
' =         ��ʂ��E�Ɉړ�                              ��ʂ��E�Ɉړ�(�X�N���[�����b�N����)
' =         �C���f���g���グ��                          �C���f���g���グ��
' =         �C���f���g��������                          �C���f���g��������
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\��              ���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉ��ړ�    ���ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ď�ړ�    ��ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����ĉE�ړ�    �E�ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         �A�N�e�B�u�Z���R�����g�̂ݕ\�����č��ړ�    ���ړ���A���Z���R�����g���g��\���h�ɂ��ăA�N�e�B�u�Z���R�����g���g�\���h�ɂ���
' =         ���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��  �A�N�e�B�u�Z���R�����g�ݒ��؂�ւ���
' =         Excel�������`�����{                         Excel�������`�����{
' =         Excel�������`������                         Excel�������`������
' =         �Z���R�����g�̏����ݒ���ꊇ�ύX            �Z���R�����g�̏����ݒ���ꊇ�ύX
' =         Diff�F�t��                                  �I��͈͂�Diff�`���̃t�H���g�F�ɕύX����B(��:�ԁA�V:��)
' =         �I��͈̓A�h���X����������R�s�[_XXX        �I��͈͂̃Z���A�h���X���������ĕ�����R�s�[
' =
' =     �E�I�u�W�F�N�g����
' =         �őO�ʂֈړ�                                �őO�ʂֈړ�����
' =         �Ŕw�ʂֈړ�                                �Ŕw�ʂֈړ�����
' =         �I�u�W�F�N�g�T�C�Y�ύX�v���p�e�B�ꊇ�ύX    ���݃V�[�g�̂�S�I�u�W�F�N�g��
' =                                                     �u�Z���ɍ��킹�Ĉړ��ƃT�C�Y�ύX������v�ɕύX
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
'������Mng_Clipboard.bas/GetFromClipboard()������
#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If
Private Const GHND = &H42
Private Const CF_TEXT = &H1
Private Const CF_LINK = &HBF00
Private Const CF_BITMAP = 2
Private Const CF_METAFILE = 3
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const MAXSIZE = 4096
'������Mng_Clipboard.bas/GetFromClipboard()������
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

'������Macro.bas/ReadSettingFile()/WriteSettingFile()������
Const sDELIMITER_INIT As String = vbTab
'������Macro.bas/ReadSettingFile()/WriteSettingFile()������

'������Mng_SysCmd.bas/ExecDosCmdRunas()������
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long
'������Mng_SysCmd.bas/ExecDosCmdRunas()������

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
    Const sFILEEXPORT_CHAR_SET As String = "Shift_JIS" '(UTF-8|UTF-16|Shift_JIS|EUC-JP|ISO-2022-JP|...)
    Const lFILEEXPORT_LINE_SEPARATER As Long = 10 '13:CR 10:LF -1:CRLF
'=== DOS�R�}���h���ꊇ���s() ===
    Const sCMDEXEBAT_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeBat.log"
    Const sCMDEXEBAT_BAT_FILE_NAME As String = "MyExcelAddinCmdExeBat.bat"
    Const bCMDEXEBAT_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOS�R�}���h���ꊇ���s_�Ǘ��Ҍ���() ===
    Const sCMDEXEBATRUNAS_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeBatRunas.log"
    Const bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOS�R�}���h���e�X���s() ===
    Const sCMDEXEUNI_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeUni.log"
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
'=== �I��͈̓A�h���X����������R�s�[_xxx() ===
    Const sCELLADRJOIN_DELIMITER As String = ""
    Const bCELLADRJOIN_FORMAT_R1C1 As Boolean = False
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
    Dim bCmntVsblEnb As Boolean
    bCmntVsblEnb = ReadSettingFile("bCMNT_VSBL_ENB", bCMNT_VSBL_ENB)
    
    '*** �V���[�g�J�b�g�L�[�ݒ�X�V ***
    ' <<�V���[�g�J�b�g�L�[�ǉ����@>>
    '   dMacroShortcutKeys�ɑ΂��ăL�[<�}�N����>�A�l<�V���[�g�J�b�g�L�[>��ǉ�����B
    '   �������ɂ̓V���[�g�J�b�g�L�[�A�������Ƀ}�N�������w�肷��B
    '   �V���[�g�J�b�g�L�[�� Ctrl �� Shift �ȂǂƑg�ݍ��킹�Ďw��ł���B
    '     Ctrl�F^�AShift�F+�AAlt�F%
    '   �ڍׂ͈ȉ� URL �Q�ƁB
    '     https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '������ �ݒ� ������
    '����
'   dMacroShortcutKeys.Add "", "F1�w���v������"
    
    '�}�N���ݒ�
'   dMacroShortcutKeys.Add "", "�}�N���V���[�g�J�b�g�L�[�S�ėL����"
'   dMacroShortcutKeys.Add "", "�}�N���V���[�g�J�b�g�L�[�S�Ė�����"
    dMacroShortcutKeys.Add "+%{F8}", "�A�h�C���}�N�����s"
'   dMacroShortcutKeys.Add "", "���W���[���ꊇ�G�N�X�|�[�g_�A�h�C��"
'   dMacroShortcutKeys.Add "", "���W���[���ꊇ�G�N�X�|�[�g_�A�N�e�B�u�u�b�N"
    dMacroShortcutKeys.Add "^+f", "CtrlShiftF�}�N��"
    
    '�u�b�N����
'   dMacroShortcutKeys.Add "", "�ʃv���Z�X�ŊJ��"
    dMacroShortcutKeys.Add "^%p", "�t�@�C���p�X�R�s�["
    dMacroShortcutKeys.Add "^%n", "�t�@�C�����R�s�["
    
    '�V�[�g����
'   dMacroShortcutKeys.Add "", "EpTree�̊֐��c���[��Excel�Ŏ�荞��"
    dMacroShortcutKeys.Add "^%h", "Excel���ᎆ"
'   dMacroShortcutKeys.Add "", "�I���V�[�g�؂�o��"
    dMacroShortcutKeys.Add "^%c", "�S�V�[�g�����R�s�["
'   dMacroShortcutKeys.Add "", "�V�[�g�\����\����؂�ւ�"
'   dMacroShortcutKeys.Add "", "�V�[�g���בւ���Ɨp�V�[�g���쐬"
    dMacroShortcutKeys.Add "^%{PGUP}", "�V�[�g�I���E�B���h�E��\��"
    dMacroShortcutKeys.Add "^%{PGDN}", "�V�[�g�I���E�B���h�E��\��"
'   dMacroShortcutKeys.Add "", "�V�[�g���ꊇ�ύX"
    dMacroShortcutKeys.Add "+{F11}", "�V�[�g�ǉ��J�X�^��"
    dMacroShortcutKeys.Add "^%{HOME}", "�擪�V�[�g�փW�����v"
    dMacroShortcutKeys.Add "^%{END}", "�����V�[�g�փW�����v"
'   dMacroShortcutKeys.Add "", "�V�[�g�Čv�Z���Ԍv��"
    
    '�Z������
'   dMacroShortcutKeys.Add "", "�t�@�C���G�N�X�|�[�g"
'   dMacroShortcutKeys.Add "", "DOS�R�}���h���ꊇ���s"
'   dMacroShortcutKeys.Add "", "DOS�R�}���h���e�X���s"
'   dMacroShortcutKeys.Add "", "DOS�R�}���h���ꊇ���s_�Ǘ��Ҍ���"
'   dMacroShortcutKeys.Add "^+f", "���������̕����F��ύX" '�uCtrlShiftF�}�N���v�ɂĎ��s
'   dMacroShortcutKeys.Add "", "�Z�����̊ې������f�N�������g"
'   dMacroShortcutKeys.Add "", "�Z�����̊ې������C���N�������g"
'   dMacroShortcutKeys.Add "", "�c���[���O���[�v��"
'   dMacroShortcutKeys.Add "", "�I��͈͂̃Z���A�h���X���������ĕ�����R�s�["
'   dMacroShortcutKeys.Add "", "�n�C�p�[�����N�ꊇ�I�[�v��"
    dMacroShortcutKeys.Add "^+j", "�n�C�p�[�����N�Ŕ��"
'   dMacroShortcutKeys.Add "", "�I��͈͓��Œ���"
    dMacroShortcutKeys.Add "^+c", "�͈͂��ێ������܂܃Z���R�s�["
    dMacroShortcutKeys.Add "^+d", "��s�ɂ܂Ƃ߂ăZ���R�s�["
    dMacroShortcutKeys.Add "^%d", "���ݒ�ύX����s�ɂ܂Ƃ߂ăZ���R�s�["
'   dMacroShortcutKeys.Add "^+v", "�N���b�v�{�[�h�l�\��t��" '�}�N���g�p���̓A���h�D�ł��Ȃ����߁A�ɗ͎g�p���Ȃ�
    dMacroShortcutKeys.Add "^2", "�w�i�F���g�O��"
    dMacroShortcutKeys.Add "^%2", "���ݒ�ύX���w�i�F���g�O���̐F�I��"
    dMacroShortcutKeys.Add "+%2", "���ݒ�ύX���w�i�F���g�O���̐F�X�|�C�g"
    dMacroShortcutKeys.Add "^3", "�t�H���g�F���g�O��"
    dMacroShortcutKeys.Add "^%3", "���ݒ�ύX���t�H���g�F���g�O���̐F�I��"
    dMacroShortcutKeys.Add "+%3", "���ݒ�ύX���t�H���g�F���g�O���̐F�X�|�C�g"
'   dMacroShortcutKeys.Add "^%{DOWN}", "'�I�[�g�t�B�����s(""Down"")'"
'   dMacroShortcutKeys.Add "^%{UP}", "'�I�[�g�t�B�����s(""Up"")'"
    dMacroShortcutKeys.Add "^%{UP}", "��ʂ���Ɉړ�"
    dMacroShortcutKeys.Add "^%{DOWN}", "��ʂ����Ɉړ�"
    dMacroShortcutKeys.Add "^%{LEFT}", "��ʂ����Ɉړ�"
    dMacroShortcutKeys.Add "^%{RIGHT}", "��ʂ��E�Ɉړ�"
    dMacroShortcutKeys.Add "^+>", "�C���f���g���グ��"
    dMacroShortcutKeys.Add "^+<", "�C���f���g��������"
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
    dMacroShortcutKeys.Add "^+{F11}", "���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��"
    dMacroShortcutKeys.Add "^+i", "Excel�������`�����{"
    dMacroShortcutKeys.Add "^%i", "Excel�������`������"
'   dMacroShortcutKeys.Add "", "�Z���R�����g�̏����ݒ���ꊇ�ύX"
    dMacroShortcutKeys.Add "+%d", "Diff�F�t��"
    
    '�I�u�W�F�N�g����
'   dMacroShortcutKeys.Add "^+f", "�őO�ʂֈړ�" '�uCtrlShiftF�}�N���v�ɂĎ��s
    dMacroShortcutKeys.Add "^+b", "�Ŕw�ʂֈړ�"
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
Private Sub �����������O�����J�p�}�N������������()
    '�v���V�[�W�����X�g�\���p�̃_�~�[�v���V�[�W��
End Sub

' ������ ���� ������
' =============================================================================
' = �T�v    F1�w���v�𖳌�������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub F1�w���v������()
    Application.OnKey "{F1}", ""
End Sub

' ������ �}�N���ݒ� ������
' =============================================================================
' = �T�v    �}�N���V���[�g�J�b�g�L�[�S�ėL����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub �}�N���V���[�g�J�b�g�L�[�S�ėL����()
    Call SwitchMacroShortcutKeysActivation(True)
    
    Application.StatusBar = "�������}�N���V���[�g�J�b�g�L�[��L�������܂���������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v    �}�N���V���[�g�J�b�g�L�[�S�Ė�����
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub �}�N���V���[�g�J�b�g�L�[�S�Ė�����()
    Call SwitchMacroShortcutKeysActivation(False)
    
    Application.StatusBar = "�������}�N���V���[�g�J�b�g�L�[�𖳌������܂���������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
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

' =============================================================================
' = �T�v    �{�A�h�C�����̑S�}�N��/�v���V�[�W�����G�N�X�|�[�g����
' = �o��    �E�ȉ��̎Q�Ɛݒ��ǉ�����K�v����B
' =           - [�c�[��] -> [�Q�Ɛݒ�] ->�uMicrosoft Visual Basic for Applications Extensibility�v
' = �ˑ�    Macros.bas/ExportAllModules()
' = ����    Macros.bas
' =============================================================================
Public Sub ���W���[���ꊇ�G�N�X�|�[�g_�A�h�C��()
    Const sMACRO_NAME As String = "���W���[���ꊇ�G�N�X�|�[�g_�A�h�C��"
    Call ExportAllModules(ThisWorkbook)
    MsgBox "�A�h�C�����̑S���W���[�����G�N�X�|�[�g���܂����I", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N���̑S�}�N��/�v���V�[�W�����G�N�X�|�[�g����
' = �o��    �E�ȉ��̎Q�Ɛݒ��ǉ�����K�v����B
' =           - [�c�[��] -> [�Q�Ɛݒ�] ->�uMicrosoft Visual Basic for Applications Extensibility�v
' = �ˑ�    Macros.bas/ExportAllModules()
' = ����    Macros.bas
' =============================================================================
Public Sub ���W���[���ꊇ�G�N�X�|�[�g_�A�N�e�B�u�u�b�N()
    Const sMACRO_NAME As String = "���W���[���ꊇ�G�N�X�|�[�g_�A�N�e�B�u�u�b�N"
    Call ExportAllModules(ActiveWorkbook)
    MsgBox "�A�N�e�B�u�u�b�N���̑S���W���[�����G�N�X�|�[�g���܂����I", vbOKOnly, sMACRO_NAME
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

' ������ �u�b�N���� ������
' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N��ʃv���Z�X�ŊJ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �ʃv���Z�X�ŊJ��()
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sActiveBookPath
    sActiveBookPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    objWshShell.Run "cmd /c excel /x /r """ & sActiveBookPath & """", 0, False
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N�̃t�@�C���p�X���R�s�[
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �t�@�C���p�X�R�s�[()
    Const sMACRO_NAME As String = "�t�@�C���p�X�R�s�["
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
        .PutInClipboard
    End With
    
    '*** �t�B�[�h�o�b�N ***
    Application.StatusBar = "���������������� " & sMACRO_NAME & "�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' =============================================================================
' = �T�v    �A�N�e�B�u�u�b�N�̃t�@�C�������R�s�[
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �t�@�C�����R�s�[()
    Const sMACRO_NAME As String = "�t�@�C�����R�s�["
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText ActiveWorkbook.Name
        .PutInClipboard
    End With
    
    '*** �t�B�[�h�o�b�N ***
    Application.StatusBar = "���������������� " & sMACRO_NAME & "�����I ����������������"
    Sleep 200 'ms �P��
    Application.StatusBar = False
End Sub

' ������ �V�[�g���� ������
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
    sOutSheetName = ReadSettingFile("sEPTREE_OUT_SHEET_NAME", sEPTREE_OUT_SHEET_NAME)
    lMaxFuncLevelIni = ReadSettingFile("lEPTREE_MAX_FUNC_LEVEL_INI", lEPTREE_MAX_FUNC_LEVEL_INI)
    lClmWidth = ReadSettingFile("lEPTREE_CLM_WIDTH", lEPTREE_CLM_WIDTH)
    
    'Eptree���O�t�@�C���p�X�擾
    sEptreeLogPath = ReadSettingFile("sEPTREE_OUT_LOG_PATH", sEPTREE_OUT_LOG_PATH)
    sEptreeLogPath = ShowFileSelectDialog(sEptreeLogPath, "EpTreeLog.txt�̃t�@�C���p�X��I�����Ă�������")
    If sEptreeLogPath = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call WriteSettingFile("sEPTREE_OUT_LOG_PATH", sEptreeLogPath)
    
    '�J���p���[�g�t�H���_�擾
    sDevRootDirPath = ReadSettingFile("sEPTREE_DEV_ROOT_DIR_PATH", sEPTREE_DEV_ROOT_DIR_PATH)
    sDevRootDirPath = ShowFolderSelectDialog(sDevRootDirPath, "�J���p���[�g�t�H���_�p�X��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j")
    If sDevRootDirPath = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    sDevRootDirName = ExtractTailWord(sDevRootDirPath, "\")
    Call WriteSettingFile("sEPTREE_DEV_ROOT_DIR_PATH", sDevRootDirPath)
    
    '���[�g�t�H���_���x���擾
    lDevRootLevel = ReadSettingFile("lEPTREE_DEV_ROOT_DIR_LEVEL", lEPTREE_DEV_ROOT_DIR_LEVEL)
    Dim sDevRootLevel As String
    sDevRootLevel = InputBox("���[�g�t�H���_���x������͂��Ă�������", sMACRO_NAME, CStr(lDevRootLevel))
    If sDevRootLevel = "" Then
        MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call WriteSettingFile("lEPTREE_DEV_ROOT_DIR_LEVEL", sDevRootLevel)
    
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
' = �T�v    Excel���ᎆ
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub Excel���ᎆ()
    '�A�h�C���ݒ�ǂݏo��
    Dim sFontName As String
    Dim lFontSize As Long
    Dim lClmWidth As Long
    sFontName = ReadSettingFile("sEXCELGRID_FONT_NAME", sEXCELGRID_FONT_NAME)
    lFontSize = ReadSettingFile("lEXCELGRID_FONT_SIZE", lEXCELGRID_FONT_SIZE)
    lClmWidth = ReadSettingFile("lEXCELGRID_CLM_WIDTH", lEXCELGRID_CLM_WIDTH)
    
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

' =============================================================================
' = �T�v    �V�[�g�I���E�B���h�E��\������
' = �o��    �E���ɂ���Ă̓V�[�g���A�N�e�B�u������Ȃ����Ƃ����邪�A
' =           �Ȃ������O��MsgBox����ΑΏ��ł���B
' =         �EbUseMyUserForm = True�ɂ��A���샆�[�U�t�H�[���̃V�[�g
' =           �I���E�B���h�E��\���ł���
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �V�[�g�I���E�B���h�E��\��()
    Const bUseMyUserForm As Boolean = True
    If bUseMyUserForm = True Then
        SelectActivationSheet.Show
    Else
        Dim bMsgBoxShow As Boolean
        bMsgBoxShow = ReadSettingFile("bSHTSELWIN_MSGBOX_SHOW", bSHTSELWIN_MSGBOX_SHOW)
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
    End If
End Sub

' =============================================================================
' = �T�v    �V�[�g�����ꊇ�ύX����
' = �o��    �E���v�����F2�s�ڈȍ~�A2��ڂɋ��V�[�g���A3��ڂɐV�V�[�g�����w�肷��B
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �V�[�g���ꊇ�ύX()
    Const lOLD_SHTNAME_CLM As Long = 2
    Const lNEW_SHTNAME_CLM As Long = 3
    Const lSTART_ROW As Long = 2
    Application.ScreenUpdating = False
    With ActiveSheet
        Dim lStrtRow As Long
        Dim lLastRow As Long
        lStrtRow = lSTART_ROW
        lLastRow = .Cells(.Rows.Count, lOLD_SHTNAME_CLM).End(xlUp).Row
        Dim lRowIdx As Long
        For lRowIdx = lStrtRow To lLastRow
            Dim sShtNameOld As String
            Dim sShtNameNew As String
            sShtNameOld = .Cells(lRowIdx, lOLD_SHTNAME_CLM).Value
            sShtNameNew = .Cells(lRowIdx, lNEW_SHTNAME_CLM).Value
            If sShtNameOld <> sShtNameNew Then
                ActiveWorkbook.Sheets(sShtNameOld).Name = sShtNameNew
            End If
        Next
    End With
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �V�[�g��ǉ�����i�J�X�^���ݒ�Łj
' = �o��    �E�V�[�g�ǉ����A�ȉ������{����
' =           - �A�E�g���C�����ɏW�v�s����A�W�v������ɐݒ肷��
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �V�[�g�ǉ��J�X�^��()
    'MsgBox "�J�X�^���ݒ�ŃV�[�g�ǉ�"
    Application.ScreenUpdating = False
    Dim shAddSht As Worksheet
    Set shAddSht = ActiveWorkbook.Sheets.Add()
    shAddSht.Outline.SummaryRow = xlAbove
    shAddSht.Outline.SummaryColumn = xlLeft
    Application.ScreenUpdating = True
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

' ==================================================================
' = �T�v    �V�[�g���ɍČv�Z�ɂ����鎞�Ԃ��v������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Public Sub �V�[�g�Čv�Z���Ԍv��()
    Const lLOOP_NUM As Long = 10
    
    Application.ScreenUpdating = False
    
    Dim previonsCalculationcMode As Variant
    previonsCalculationcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' ���͒l��1�����̏ꍇ�͏I������
    If Not (lLOOP_NUM > 0) Then
        Exit Sub
    End If
    
    ' ### �x���`�}�[�N�J�n ###
    ' �w�b�_�[(�V�[�g���̈ꗗ)���o��
    Dim shTrgtSheet As Worksheet
    For Each shTrgtSheet In ActiveWorkbook.Worksheets
        Debug.Print shTrgtSheet.Name & vbTab;
    Next
    Debug.Print
    
    Dim lLoopIdx As Long
    For lLoopIdx = 1 To lLOOP_NUM
        ' �V�[�g���ƂɍČv�Z&�������Ԃ��o��
        Dim vStartTime As Variant
        Dim vFinishTime As Variant
        For Each shTrgtSheet In Worksheets
            shTrgtSheet.Cells.Dirty
            vStartTime = Timer
            shTrgtSheet.Calculate
            vFinishTime = Timer
            Debug.Print Format(vFinishTime - vStartTime, "0.0000") & vbTab;
        Next
        Debug.Print
    Next
    
    Application.Calculation = previonsCalculationcMode
    Application.ScreenUpdating = True
End Sub

' ������ �Z������ ������
' =============================================================================
' = �T�v    �I��͈͂��t�@�C���Ƃ��ăG�N�X�|�[�g����B
' =         �ׂ荇������̃Z���ɂ̓^�u������}�����ďo�͂���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_FileSys.bas/OutputTxtFile()
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
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bFILEEXPORT_IGNORE_INVISIBLE_CELL", bFILEEXPORT_IGNORE_INVISIBLE_CELL)
    
    '*** �o�͐���� ***
    '�t�H���_�p�X
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputDirPathInit As String
    Dim sOutputDirPath As String
    sOutputDirPathInit = objWshShell.SpecialFolders("Desktop")
    sOutputDirPath = ReadSettingFile("sFILEEXPORT_OUT_DIR_PATH", sOutputDirPathInit)
    sOutputDirPath = ShowFolderSelectDialog(sOutputDirPath)
    If sOutputDirPath = "" Then
        MsgBox "�����ȃt�H���_���w��������̓t�H���_���I������܂���ł����B", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        End
    Else
        'Do Nothing
    End If
    Call WriteSettingFile("sFILEEXPORT_OUT_DIR_PATH", sOutputDirPath)
    
    '�t�@�C����
    Dim sOutputFileName As String
    Dim sOutputFilePath As String
    Dim sFileExt As String
    Dim sDelimiter As String
    sOutputFileName = ReadSettingFile("sFILEEXPORT_OUT_FILE_NAME", sFILEEXPORT_OUT_FILE_NAME)
    sOutputFileName = InputBox("�t�@�C��������͂��Ă��������B(�g���q�t��)", sMACRO_NAME, sOutputFileName)
    If InStr(sOutputFileName, ".") Then
        'Do Nothing
    Else
        MsgBox "�t�@�C�������w�肳��܂���ł����B", vbCritical, sMACRO_NAME
        MsgBox "�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        End
    End If
    Call WriteSettingFile("sFILEEXPORT_OUT_FILE_NAME", sOutputFileName)
    
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
    
    '*** Range�^����String()�^�֕ϊ� ***
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                sDelimiter _
            )
    
    '*** �t�@�C���o�͏��� ***
    Call OutputTxtFile(sOutputFilePath, asRange, sFILEEXPORT_CHAR_SET, lFILEEXPORT_LINE_SEPARATER)
    
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
' = �o��    �E��ʂ̃R�}���h�����s����ꍇ�A�uDOS�R�}���h���e�X���s()�v�ɔ�ׂ�
' =           �{�}�N���̂ق��������B
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub DOS�R�}���h���ꊇ���s()
    Const sMACRO_NAME As String = "DOS�R�}���h���ꊇ���s"
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEBAT_IGNORE_INVISIBLE_CELL", bCMDEXEBAT_IGNORE_INVISIBLE_CELL)
    
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
    Print #1, ""
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Close #1
    Call ExecDosCmd(sBatFilePath & " >> " & sOutputFilePath, False)
    
    '*** �o�b�`�t�@�C���폜 ***
    Kill sBatFilePath
    
    MsgBox "���s�����I", vbOKOnly, sMACRO_NAME
    
    '*** �o�̓t�@�C�����J�� ***
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v    �I��͈͓���DOS�R�}���h���o�b�`�t�@�C���ɏ����o���Ă܂Ƃ߂Ď��s����B�i�Ǘ��Ҍ����j
' =         �P���I�����̂ݗL���B
' = �o��    �Ȃ�
' = �ˑ�    Mng_SysCmd.bas/ExecDosCmdRunas()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub DOS�R�}���h���ꊇ���s_�Ǘ��Ҍ���()
    Const sMACRO_NAME As String = "DOS�R�}���h���ꊇ���s_�Ǘ��Ҍ���"
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL", bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL)
    
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
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEBATRUNAS_REDIRECT_FILE_NAME
    
    '*** �R�}���h���s ***
    Open sOutputFilePath For Append As #1
    Print #1, ""
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Print #1, ExecDosCmdRunas(asRange, True)
    Close #1
    
    MsgBox "���s�����I", vbOKOnly, sMACRO_NAME
    
    '*** �o�̓t�@�C�����J�� ***
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = �T�v    �I��͈͓���DOS�R�}���h�����ꂼ����s����B
' =         �P���I�����̂ݗL���B
' = �o��    �E�P���̃R�}���h�����s����ꍇ�A�uDOS�R�}���h���ꊇ���s()�v�ɔ�ׂ�
' =           �{�}�N���̂ق��������B
' =         �E��ʂ̃R�}���h�����s����ہA�R�}���h���Ƀv�����v�g���\�������B
' =           �ڏ��Ɋ�����ꍇ�́A�uDOS�R�}���h���ꊇ���s()�v�����s���邱�ƁB
' = �ˑ�    Mng_Array.bas/ConvRange2Array()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub DOS�R�}���h���e�X���s()
    Const sMACRO_NAME As String = "DOS�R�}���h���e�X���s"
    
    '*** �A�h�C���ݒ�ǂݏo�� ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEUNI_IGNORE_INVISIBLE_CELL", bCMDEXEUNI_IGNORE_INVISIBLE_CELL)
    
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
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
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
    Const lREGEXP_IGNORECASE As Boolean = False
    
    Dim cCLR_RGBS As Variant
    Set cCLR_RGBS = CreateObject("System.Collections.ArrayList")
    '�������F�ݒ聥����
    Const sCOLOR_TYPE As String = "0:�ԁA1:���A2:�΁A3:���A4:��A5:���A6:���A7:��"
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
    Dim sSrchStr As String
    Dim lClrRgbInit As Long
    sSrchStr = ReadSettingFile("sWORDCOLOR_SRCH_WORD", sWORDCOLOR_SRCH_WORD)
    lClrRgbInit = ReadSettingFile("lWORDCOLOR_CLR_RGB", lWORDCOLOR_CLR_RGB)
    
    '�����Ώە�����I��
    sSrchStr = InputBox("����������𐳋K�\���œ��͂��Ă�������", sMACRO_NAME, sSrchStr)
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
    Call WriteSettingFile("sWORDCOLOR_SRCH_WORD", sSrchStr)
    Call WriteSettingFile("lWORDCOLOR_CLR_RGB", lClrRgbSelected)
    
    '�Ώ۔͈͓���(�I��͈͂Ǝg�p����Ă���͈͂̋��ʕ���)
    Dim rTrgtRng As Range
    Set rTrgtRng = Application.Intersect(Selection, ActiveSheet.UsedRange)
    
    '����������F�ύX
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSrchStr
    oRegExp.IgnoreCase = lREGEXP_IGNORECASE
    oRegExp.Global = True
    Dim oMatchResult
    Dim oCell As Range
    For Each oCell In rTrgtRng
        If oCell.Value <> "" Then
            Dim sTargetStr
            sTargetStr = oCell.Value
            Set oMatchResult = oRegExp.Execute(sTargetStr)
            Dim lMatchIdx As Long
            For lMatchIdx = 0 To oMatchResult.Count - 1
                Dim lCharPos As Long
                lCharPos = oMatchResult(lMatchIdx).FirstIndex + 1
                oCell.Characters( _
                    Start:=lCharPos, _
                    Length:=oMatchResult(lMatchIdx).Length _
                ).Font.Color = lClrRgbSelected
            Next lMatchIdx
        End If
    Next
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    
    MsgBox "�����I", vbOKOnly, sMACRO_NAME
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
' = �T�v    �A�N�e�B�u�Z������n�C�p�[�����N��ɔ��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �n�C�p�[�����N�Ŕ��()
    Dim rTrgtCell As Range
    On Error Resume Next
    For Each rTrgtCell In Selection
        rTrgtCell.Hyperlinks(1).Follow NewWindow:=True
        If Err.Number = 0 Then
            'Do Nothing
        Else
            Debug.Print "[" & Now & "] Error " & _
                        "[Macro] �n�C�p�[�����N�Ŕ�� " & _
                        "[Error No." & Err.Number & "] " & Err.Description
        End If
    Next
    On Error GoTo 0
End Sub

' =============================================================================
' = �T�v    �I���Z���ɑ΂��āu�I��͈͓��Œ����v�����s����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈͓��Œ���()
    If Selection(1).HorizontalAlignment = xlCenterAcrossSelection Then
        Selection.HorizontalAlignment = xlGeneral
    Else
        Selection.HorizontalAlignment = xlCenterAcrossSelection
    End If
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
    Dim bIgnoreInvisible As Boolean
    bIgnoreInvisible = ReadSettingFile("bCELLCOPYRNG_IGNORE_INVISIBLE_CELL", bCELLCOPYRNG_IGNORE_INVISIBLE_CELL)
    
    Dim sDelimiter As String
    sDelimiter = ReadSettingFile("sCELLCOPYRNG_DELIMITER", sCELLCOPYRNG_DELIMITER)
    
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
    Dim bIgnoreInvisibleCell As Boolean
    Dim bIgnoreBlankCell As Boolean
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    bIgnoreInvisibleCell = ReadSettingFile("bCELLCOPYLINE_IGNORE_INVISIBLE_CELL", bCELLCOPYLINE_IGNORE_INVISIBLE_CELL)
    bIgnoreBlankCell = ReadSettingFile("bCELLCOPYLINE_IGNORE_BLANK_CELL", bCELLCOPYLINE_IGNORE_BLANK_CELL)
    sPreffix = ReadSettingFile("sCELLCOPYLINE_PREFFIX", sCELLCOPYLINE_PREFFIX)
    sDelimiter = ReadSettingFile("sCELLCOPYLINE_DELIMITER", sCELLCOPYLINE_DELIMITER)
    sSuffix = ReadSettingFile("sCELLCOPYLINE_SUFFIX", sCELLCOPYLINE_SUFFIX)
    
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
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    sPreffix = ReadSettingFile("sCELLCOPYLINE_PREFFIX", sCELLCOPYLINE_PREFFIX)
    sDelimiter = ReadSettingFile("sCELLCOPYLINE_DELIMITER", sCELLCOPYLINE_DELIMITER)
    sSuffix = ReadSettingFile("sCELLCOPYLINE_SUFFIX", sCELLCOPYLINE_SUFFIX)
    
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
        Call WriteSettingFile("sCELLCOPYLINE_PREFFIX", sPreffix)
        Call WriteSettingFile("sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call WriteSettingFile("sCELLCOPYLINE_SUFFIX", sSuffix)
        MsgBox _
            "�ݒ��ύX���܂���" & vbNewLine & _
            "�@�擪�����F" & sPreffix & vbNewLine & _
            "�@��؂蕶���F" & sDelimiter & vbNewLine & _
            "�@���������F" & sSuffix, _
            vbOKOnly, _
            sMACRO_NAME
    ElseIf vRet = vbNo Then
        Call WriteSettingFile("sCELLCOPYLINE_PREFFIX", sPreffix)
        Call WriteSettingFile("sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call WriteSettingFile("sCELLCOPYLINE_SUFFIX", sSuffix)
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
' = �T�v    �N���b�v�{�[�h����l�\��t������
' = �o��    �E���݂̑I��͈͖͂�������
' = �ˑ�    Mng_Clipboard.bas/GetFromClipboard()
' = ����    Macro.bas
' =============================================================================
Public Sub �N���b�v�{�[�h�l�\��t��()
    Dim bResult As Boolean
    Dim sStr As String
    bResult = GetFromClipboard(sStr)
    If bResult = True Then
        ActiveSheet.PasteSpecial Format:="�e�L�X�g"
    Else
        'Do Nothing
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
    Dim lClrRgb As Long
    lClrRgb = ReadSettingFile("lCLRTGLFONT_CLR_RGB", lCLRTGLFONT_CLR_RGB)
    
    '�t�H���g�F�ύX
    If Selection(1).Font.Color = lClrRgb Then
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Font.Color = lClrRgb
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
    Dim lClrRgbInit As Long
    lClrRgbInit = ReadSettingFile("lCLRTGLFONT_CLR_RGB", lCLRTGLFONT_CLR_RGB)
    
    '�F�I��
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "�F�I�������s���܂����̂ŁA�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    '�A�h�C���ݒ�X�V
    Call WriteSettingFile("lCLRTGLFONT_CLR_RGB", lClrRgbSelected)
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
    Call WriteSettingFile("lCLRTGLFONT_CLR_RGB", lClrRgb)
End Sub

' =============================================================================
' = �T�v    �w�i�F���u�ݒ�F�v�́u�w�i�F�Ȃ��v�Ńg�O������
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' = ����    Macros.bas
' =============================================================================
Public Sub �w�i�F���g�O��()
    '�A�h�C���ݒ�ǂݏo��
    Dim lClrRgb As Long
    lClrRgb = ReadSettingFile("lCLRTGLBG_CLR_RGB", lCLRTGLBG_CLR_RGB)
    
    '�w�i�F�ύX
    If Selection(1).Interior.Color = lClrRgb Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = lClrRgb
    End If
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
    Dim lClrRgbInit As Long
    lClrRgbInit = ReadSettingFile("lCLRTGLBG_CLR_RGB", lCLRTGLBG_CLR_RGB)
    
    '�F�I��
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "�F�I�������s���܂����̂ŁA�����𒆒f���܂��B", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    '�A�h�C���ݒ�X�V
    Call WriteSettingFile("lCLRTGLBG_CLR_RGB", lClrRgbSelected)
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
    Call WriteSettingFile("lCLRTGLBG_CLR_RGB", lClrRgb)
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
' = �T�v    ��ʂ���Ɉړ�(�X�N���[�����b�N����)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub ��ʂ���Ɉړ�()
    With ActiveWindow
        If .ScrollRow > 1 Then
            .ScrollRow = .ScrollRow - 1
        End If
    End With
End Sub

' =============================================================================
' = �T�v    ��ʂ����Ɉړ�(�X�N���[�����b�N����)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub ��ʂ����Ɉړ�()
    With ActiveWindow
        .ScrollRow = .ScrollRow + 1
    End With
End Sub

' =============================================================================
' = �T�v    ��ʂ����Ɉړ�(�X�N���[�����b�N����)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub ��ʂ����Ɉړ�()
    With ActiveWindow
        If .ScrollColumn > 1 Then
            .ScrollColumn = .ScrollColumn - 1
        End If
    End With
End Sub

' =============================================================================
' = �T�v    ��ʂ��E�Ɉړ�(�X�N���[�����b�N����)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub ��ʂ��E�Ɉړ�()
    With ActiveWindow
        .ScrollColumn = .ScrollColumn + 1
    End With
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
' = �T�v    �A�N�e�B�u�Z���̃R�����g�\���̗L��/������؂�ւ���
' = �o��    �Ȃ�
' = �ˑ�    SettingFile.cls
' =         Macros.bas/SwitchMacroShortcutKeysActivation()
' = ����    Macros.bas
' =============================================================================
Public Sub ���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��()
    Const sMACRO_NAME As String = "���ݒ�ύX���A�N�e�B�u�Z���R�����g�̂ݕ\��"
    
    '�A�h�C���ݒ�t�@�C���ǂݏo��
    Dim bExistSetting As Boolean
    bExistSetting = ReadSettingFile("bCMNT_VSBL_ENB", bCMNT_VSBL_ENB)
    
    '�A�N�e�B�u�Z���R�����g�ݒ�X�V
    Dim bCmntVsblEnb As Boolean
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
    
    Call WriteSettingFile("bCMNT_VSBL_ENB", bCmntVsblEnb)
    
    '�V���[�g�J�b�g�L�[�ݒ� �X�V(�L����)
    Call SwitchMacroShortcutKeysActivation(True)
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
' = �T�v    ���݃V�[�g�̑S�Z���R�����g�I�u�W�F�N�g��
' =         �u�Z���ɍ��킹�Ĉړ���T�C�Y�ύX������v�Ɉꊇ�ݒ肷��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �Z���R�����g�̏����ݒ���ꊇ�ύX()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    With ActiveSheet
        Dim lLastRow As Long
        Dim lLastClm As Long
        Dim lRowIdx As Long
        Dim lClmIdx As Long
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lLastClm = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For lRowIdx = 1 To lLastRow
            For lClmIdx = 1 To lLastClm
                If .Cells(lRowIdx, lClmIdx).Comment Is Nothing Then
                    'Do Nothing
                Else
                    'MsgBox .Cells(lRowIdx, lClmIdx).Value
                    .Cells(lRowIdx, lClmIdx).Comment.Shape.Placement = xlMoveAndSize
                End If
            Next lClmIdx
        Next lRowIdx
    End With
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = �T�v    �I��͈͂�Diff�`���̃t�H���g�F�ɕύX����B(��:�ԁA�V:��)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub Diff�F�t��()
    Const bUNIFIED_MODE As Boolean = True
    Dim rCell As Range
    For Each rCell In Selection
        Dim oRegExp As Object
        Set oRegExp = CreateObject("VBScript.RegExp")
        Dim sTargetStr As String
        sTargetStr = rCell.Value
        oRegExp.IgnoreCase = True
        oRegExp.Global = True
        Dim oMatchResult As Object
        
        If bUNIFIED_MODE = True Then
            oRegExp.Pattern = "^\+"
        Else
            oRegExp.Pattern = "^>"
        End If
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Color = RGB(0, 176, 80)
        End If
        
        If bUNIFIED_MODE = True Then
            oRegExp.Pattern = "^-"
        Else
            oRegExp.Pattern = "^<"
        End If
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Color = RGB(255, 0, 0)
        End If
        
        oRegExp.Pattern = "^\$ diff"
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Bold = True
        End If
        
        oRegExp.Pattern = "^\$ git diff"
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Bold = True
        End If
    Next
End Sub

' ������ �I�u�W�F�N�g���� ������
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
' = �T�v    ���݃V�[�g�̂�S�I�u�W�F�N�g��
' =         �u�Z���ɍ��킹�Ĉړ��ƃT�C�Y�ύX������v�ɕύX
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' =============================================================================
Public Sub �I�u�W�F�N�g�T�C�Y�ύX�v���p�e�B�ꊇ�ύX()
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        objShp.Placement = xlMoveAndSize
    Next
    MsgBox "�A�N�e�B�u�V�[�g���̑S�I�u�W�F�N�g�̃v���p�e�B���A" & vbNewLine & _
        "�u�Z���ɍ��킹�Ĉړ��ƃT�C�Y�ύX������v�Ɉꊇ�ύX���܂����I"
End Sub

' =============================================================================
' = �T�v    �I��͈͂̃Z���A�h���X���������ĕ�����R�s�[
' = �o��    �Ȃ�
' = �ˑ�    Macros.bas/CopyConcatedCellAddresses()
' = ����    Macros.bas
' =============================================================================
Public Sub �I��͈̓A�h���X����������R�s�[_���΍s_���Η�()
    Call CopyConcatedCellAddresses(Selection, False, False, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub �I��͈̓A�h���X����������R�s�[_��΍s_���Η�()
    Call CopyConcatedCellAddresses(Selection, True, False, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub �I��͈̓A�h���X����������R�s�[_���΍s_��Η�()
    Call CopyConcatedCellAddresses(Selection, False, True, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub �I��͈̓A�h���X����������R�s�[_��΍s_��Η�()
    Call CopyConcatedCellAddresses(Selection, True, True, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
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
' = ����    bGetStdout  Boolean  [in]   �W���o�͎擾�L��(�ȗ���)
' = �ߒl                String          �W���o��
' = �o��    �E��ʂ̏������s��bat�����s����ꍇ�AbGetStdout��False�ɂ��邱�ƁB
' =           �R�}���h�̎��s���ʂ��K�v�ȏꍇ�́A�R�}���h�Ƀ��_�C���N�g���܂߂邱�ƁB
' =             ��jCall ExecDosCmd("xxx.bat > xxx.log", False)
' =           �y���R�z
' =           Exec�͕W���o�͂ɂ��߂�o�b�t�@�̍ő��4096�o�C�g�ł���A
' =           ����ȏ�̃f�[�^��ǂݍ��ނ�AtEndOfStream���Ɍł܂邽�߁B
' =           https://community.cybozu.dev/t/topic/181/2
' = �ˑ�    �Ȃ�
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmd( _
    ByVal sCommand As String, _
    Optional bGetStdOut As Boolean = True _
) As String
    If sCommand = "" Then
        ExecDosCmd = ""
    Else
        Dim sStdOutAll As String
        sStdOutAll = ""
        If bGetStdOut = True Then
            Dim oExeResult As Object
            Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
            Do While Not oExeResult.StdOut.AtEndOfStream
                Dim sStdOut As String
                sStdOut = oExeResult.StdOut.ReadLine
                Debug.Print sStdOut
                sStdOutAll = sStdOutAll & vbNewLine & sStdOut
            Loop
            Set oExeResult = Nothing
        Else
            Call CreateObject("WScript.Shell").Run("%ComSpec% /c """ & sCommand & """", WaitOnReturn:=True)
        End If
        ExecDosCmd = sStdOutAll
    End If
End Function

' ==================================================================
' = �T�v    �R�}���h�����s�i�Ǘ��Ҍ����j
' = ����    asCommands()    String   [in] ���s�R�}���h
' = ����    bDelFiles       Boolean  [in] Bat/Log�t�@�C���폜(�ȗ���)
' = �ߒl                    String        �W���o�́��W���G���[�o��
' = �o��    �EDesktop�t�H���_�p�X�ɋ󔒂��܂܂��ꍇ�́A���삵�Ȃ��B
' = �ˑ�    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/OutputTxtFile()
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmdRunas( _
    ByRef asCommands() As String, _
    Optional bDelFiles As Boolean = True _
) As String
    Const sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME As String = "CmdExeBatRunas"
    If Sgn(asCommands) = 0 Then
        ExecDosCmdRunas = ""
    Else
        If UBound(asCommands) < 0 Then
            ExecDosCmdRunas = ""
        Else
            Dim objWshShell
            Set objWshShell = CreateObject("WScript.Shell")
            Dim objFSO
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            
            Dim sBatFilePath As String
            Dim sLogFilePath As String
            sBatFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".bat"
            sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".log"
            
            '�u@echo off�v�}��
            ReDim Preserve asCommands(UBound(asCommands) + 1)
            Dim lIdx As Long
            For lIdx = UBound(asCommands) To (LBound(asCommands) + 1) Step -1
                asCommands(lIdx) = asCommands(lIdx - 1)
            Next lIdx
            asCommands(0) = "@echo off"
            
            'BAT�t�@�C���쐬
            Call OutputTxtFile(sBatFilePath, asCommands)
            Do While Not objFSO.FileExists(sBatFilePath)
                Sleep 100
            Loop
            
            'BAT�t�@�C�����s
            ShellExecute 0, "runas", sBatFilePath, " > " & sLogFilePath & " 2>&1", vbNullString, 1
            
            'LOG�t�@�C���o�͑҂�
            Do While Not objFSO.FileExists(sLogFilePath)
                Sleep 100
            Loop
            
            'LOG�t�@�C���Ǎ���
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(sLogFilePath, 1, True)
            Dim sStdOutAll As String
            sStdOutAll = ""
            Dim sLine As String
            Do Until objTxtFile.AtEndOfStream
                sLine = objTxtFile.ReadLine
                'MsgBox sLine
                If sStdOutAll = "" Then
                    sStdOutAll = sLine
                Else
                    sStdOutAll = sStdOutAll & vbNewLine & sLine
                End If
            Loop
            'MsgBox sStdOutAll
            objTxtFile.Close
            
            'BAT�t�@�C��/LOG�t�@�C���폜
            If bDelFiles = True Then
                Kill sBatFilePath
                Kill sLogFilePath
            End If
            
            ExecDosCmdRunas = sStdOutAll
        End If
    End If
End Function
    Private Sub Test_ExecDosCmdRunas()
        Dim asCommands() As String
        
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(0)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source2.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source2.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands, False)
    End Sub

' ============================================
' = �T�v    �z��̓��e���t�@�C���ɏ������ށB
' = ����    sFilePath       String  [in]  �o�͂���t�@�C���p�X
' =         asFileLine()    String  [in]  �o�͂���t�@�C���̓��e
' =         sCharSet        String  [in]  �����R�[�h(�ȗ���)
' =                                         (UTF-8|UTF-16|Shift_JIS|EUC-JP|ISO-2022-JP|...)
' =         lLineSeparator  Long    [in]  ���s�R�[�h(�ȗ���)
' =                                         13:CR 10:LF -1:CRLF
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ============================================
Private Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String = "shift_jis", _
    Optional ByVal lLineSeparator As Long = -1 _
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
            .LineSeparator = lLineSeparator
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
' = �T�v    �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�iWin32Api���g�p�j
' = ����    sInStr      String  [in]  �ݒ�Ώە�����
' = �ߒl                Boolean       �ݒ茋��
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
    ByVal sInStr As String _
) As Boolean
#If VBA7 Then
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr
    Dim lX As LongPtr
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim lX As Long
#End If
    Dim bResult As Boolean
    bResult = True
    
    hGlobalMemory = GlobalAlloc(GHND, LenB(sInStr) + 1)   '�ړ��\�ȃO���[�o�������������蓖��
    lpGlobalMemory = GlobalLock(hGlobalMemory)          '�u���b�N�����b�N���āA�������ւ�far�|�C���^���擾
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sInStr)      '��������O���[�o���������փR�s�[
    
    '�������̃��b�N����
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "�������̃��b�N�������ł��܂���" & vbCrLf & "���������s���܂���"
        bResult = False
    Else
        '�f�[�^���R�s�[����N���b�v�{�[�h���J��
        If OpenClipboard(0&) = 0 Then
            MsgBox "�N���b�v�{�[�h���J�����Ƃ��ł��܂���" & vbCrLf & "���������s���܂���"
            bResult = False
            Exit Function
        End If
        
        lX = EmptyClipboard()    '�N���b�v�{�[�h�̓��e������
        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory) '�f�[�^���N���b�v�{�[�h�փR�s�[
    End If
    
    '�N���b�v�{�[�h�̏�ԃ`�F�b�N
    If CloseClipboard() = 0 Then
        MsgBox "�N���b�v�{�[�h����邱�Ƃ��ł��܂���"
        bResult = False
    End If
    SetToClipboard = bResult
End Function

' ==================================================================
' = �T�v    �N���b�v�{�[�h����e�L�X�g���擾�iWin32Api���g�p�j
' = ����    sOutStr     String  [Out]   �擾�敶����
' = �ߒl                Boolean         �擾����
' = �o��    Win32API���g�p����B
' =         �� �N���b�v�{�[�h�� DataObject �� PutInClipboard �ł����p
' =            �\��������DataObject �͎Q�Ɛݒ肪�K�v�Ȃ��������̃N
' =            ���b�v�{�[�h�`���ɂ͓\��t������Ȃ���iCF_UNICODETEXT
' =            �݂̂� CF_TEXT�ւ͓\��t������Ȃ��j
' =            ��L�̂悤�� DataObject ���g�p�������Ȃ��ꍇ�ɖ{�֐�
' =            �𗘗p���邱�ơ
' = �ˑ�    user32/OpenClipboard()
' =         user32/CloseClipboard()
' =         user32/GetClipboardData()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = ����    Mng_Clipboard.bas
' ==================================================================
Public Function GetFromClipboard( _
    ByRef sOutStr As String _
) As Boolean
#If VBA7 Then
    Dim hClipMemory As LongPtr
    Dim lpClipMemory As LongPtr
#Else
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
#End If
    Dim sStr As String
    Dim lRetVal As Long
    Dim bResult As Boolean
    bResult = True
    sOutStr = ""
    
    If OpenClipboard(0&) = 0 Then
        MsgBox "�N���b�v�{�[�h���J�����Ƃ��ł��܂���" & vbCrLf & "���������s���܂���"
        bResult = False
        Exit Function
    End If
    
    ' Obtain the handle to the global memory block that is referencing the text.
    hClipMemory = GetClipboardData(CF_TEXT)
    If IsNull(hClipMemory) Then
        MsgBox "Could not allocate memory"
        bResult = False
    Else
        ' Lock Clipboard memory so we can reference the actual data string.
        lpClipMemory = GlobalLock(hClipMemory)
        
        If Not IsNull(lpClipMemory) Then
            sStr = Space$(MAXSIZE)
            Call lstrcpy(sStr, lpClipMemory)
            Call GlobalUnlock(hClipMemory)
            sStr = Mid(sStr, 1, InStr(1, sStr, Chr$(0), 0) - 1)
        Else
            MsgBox "Could not lock memory to copy string from."
            bResult = False
        End If
    End If
    
    If CloseClipboard() = 0 Then
        MsgBox "�N���b�v�{�[�h����邱�Ƃ��ł��܂���"
        bResult = False
    Else
        sOutStr = sStr
    End If
    GetFromClipboard = bResult
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
    Optional ByVal lIndentWidth As Long = 2 _
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

' ==================================================================
' = �T�v    �ݒ�t�@�C������ݒ���擾����
' = ����    sKey            String      [in]    �ݒ�L�[
' = ����    vInitValue      Variant     [in]    �ݒ�l(�����l)
' = �ߒl                    Variant             �ݒ�l
' = �o��    �E�t�@�C���I�[�v����A�ݒ�l���擾����B
' =           �ݒ�l�����݂��Ȃ��ꍇ�A�ݒ�l(�����l)�Őݒ�l���X�V���ĕۑ�����B
' =         �E�ȉ��̏ꍇ�AvInitValue��ԋp����
' =           - sFilePath�����݂��Ȃ�
' =           - sKey�����݂��Ȃ�
' = �ˑ�    Macros.bas/GetAddinSettingFilePath()
' =         Mng_FileSys.bas/CreateDirectry()
' =         Macros.bas/ConvCtrlchr2Str()
' =         Macros.bas/ConvStr2Ctrlchr()
' = ����    Macros.bas
' ==================================================================
Public Function ReadSettingFile( _
    ByVal sKey As String, _
    ByVal vInitValue As Variant _
) As Variant
    Dim dSettingItems As Object
    Set dSettingItems = CreateObject("Scripting.Dictionary")
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim vKey As Variant
    
    '�ݒ�t�@�C���p�X�擾
    Dim sFilePath As String
    sFilePath = GetAddinSettingFilePath()
    
    '�ݒ�t�@�C���ǂݏo��
    If objFSO.FileExists(sFilePath) Then
        '�t�@�C���ǂݏo��
        Open sFilePath For Input As #1
        Do Until EOF(1)
            Dim vKeyValue As Variant
            Dim sLine As String
            Line Input #1, sLine
            If InStr(sLine, sDELIMITER_INIT) Then
                vKeyValue = Split(sLine, sDELIMITER_INIT)
                If UBound(vKeyValue) = 0 Then
                    dSettingItems.Add vKeyValue(0), ""           '�P���؂蕶��(�l�Ȃ�)
                ElseIf UBound(vKeyValue) = 1 Then
                    dSettingItems.Add vKeyValue(0), vKeyValue(1) '�P���؂蕶��(�l����)
                Else
                    Stop                                          '������؂蕶��
                End If
            Else
                Stop                                              '��؂蕶���Ȃ�
            End If
        Loop
        Close #1
        
        '�ݒ荀�ڎ擾���X�V
        If dSettingItems.Exists(sKey) = True Then
            Dim sItem As String
            sItem = dSettingItems.Item(sKey)
            '�^�ϊ�
            Dim vOutValue As Variant
            Select Case VarType(vInitValue)
                Case vbInteger: vOutValue = CInt(sItem)
                Case vbLong: vOutValue = CLng(sItem)
                Case vbSingle: vOutValue = CSng(sItem)
                Case vbDouble: vOutValue = CDbl(sItem)
                Case vbBoolean: vOutValue = CBool(sItem)
                Case vbString: vOutValue = CStr(ConvStr2Ctrlchr(sItem))
                Case vbCurrency: vOutValue = CCur(sItem)
                Case vbByte: vOutValue = CByte(sItem)
                Case vbDate: vOutValue = CDate(sItem)
                Case vbVariant: vOutValue = CVar(sItem)
               'Case vbEmpty      : vOutValue = CXxx(sItem)
               'Case vbNull       : vOutValue = CXxx(sItem)
               'Case vbObject     : vOutValue = CXxx(sItem)
               'Case vbError      : vOutValue = CXxx(sItem)
               'Case vbDataObject : vOutValue = CXxx(sItem)
               'Case vbArray      : vOutValue = CXxx(sItem)
                Case Else: vOutValue = ""
            End Select
            ReadSettingFile = vOutValue
        Else
            '���ڒǉ�
            dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vInitValue))
            
            '�t�@�C���ۑ�
            Open sFilePath For Output As #1
            For Each vKey In dSettingItems
                Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
            Next
            Close #1
            ReadSettingFile = vInitValue
        End If
    Else
        '�i�[��t�H���_�쐬
        Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
        
        '���ڒǉ�
        dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vInitValue))
        
        '�t�@�C���ۑ�
        Open sFilePath For Output As #1
        For Each vKey In dSettingItems
            Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
        Next
        Close #1
        
        ReadSettingFile = vInitValue
    End If
End Function

' ==================================================================
' = �T�v    �ݒ�t�@�C������ݒ���X�V���ĕۑ�����
' = ����    sKey            String      [in]    �ݒ�L�[
' = ����    vValue          Variant     [in]    �ݒ�l
' = �ߒl                                        �Ȃ�
' = �o��    �E�t�@�C���I�[�v����A�ݒ�l���X�V/�ǉ�����B
' = �ˑ�    Macros.bas/GetAddinSettingFilePath()
' =         Mng_FileSys.bas/CreateDirectry()
' =         Macros.bas/ConvCtrlchr2Str()
' = ����    Macros.bas
' ==================================================================
Public Function WriteSettingFile( _
    ByVal sKey As String, _
    ByVal vValue As Variant _
)
    Dim dSettingItems As Object
    Set dSettingItems = CreateObject("Scripting.Dictionary")
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim vKey As Variant
    
    '�ݒ�t�@�C���p�X�擾
    Dim sFilePath As String
    sFilePath = GetAddinSettingFilePath()
    
    '�i�[��t�H���_�쐬
    Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
    
    '�ݒ�t�@�C���ǂݏo��
    Open sFilePath For Input As #1
    Do Until EOF(1)
        Dim vKeyValue As Variant
        Dim sLine As String
        Line Input #1, sLine
        If InStr(sLine, sDELIMITER_INIT) Then
            vKeyValue = Split(sLine, sDELIMITER_INIT)
            If UBound(vKeyValue) = 0 Then
                dSettingItems.Add vKeyValue(0), ""           '�P���؂蕶��(�l�Ȃ�)
            ElseIf UBound(vKeyValue) = 1 Then
                dSettingItems.Add vKeyValue(0), vKeyValue(1) '�P���؂蕶��(�l����)
            Else
                Stop                                          '������؂蕶��
            End If
        Else
            Stop                                              '��؂蕶���Ȃ�
        End If
    Loop
    Close #1
    
    '���ڒǉ�
    If dSettingItems.Exists(sKey) Then
        dSettingItems.Item(sKey) = vValue
    Else
        dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vValue))
    End If
    
    '�t�@�C���ۑ�
    Open sFilePath For Output As #1
    For Each vKey In dSettingItems
        Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
    Next
    Close #1
End Function

' ==================================================================
' = �T�v    �ݒ�l�ϊ��p ���䕶��to������ �ϊ�
' = ����    sValue          String      [in]    �l(���䕶��)
' = �ߒl                    String              �l(������)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Private Function ConvCtrlchr2Str( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case vbTab:     ConvCtrlchr2Str = "vbTab"
        Case vbCr:      ConvCtrlchr2Str = "vbCr"
        Case vbLf:      ConvCtrlchr2Str = "vbLf"
        Case vbNewLine: ConvCtrlchr2Str = "vbNewLine"
        Case Else:      ConvCtrlchr2Str = sValue
    End Select
End Function

' ==================================================================
' = �T�v    �ݒ�l�ϊ��p ������to���䕶�� �ϊ�
' = ����    sValue          String      [in]    �l(������)
' = �ߒl                    String              �l(���䕶��)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Private Function ConvStr2Ctrlchr( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case "vbTab":     ConvStr2Ctrlchr = vbTab
        Case "vbCr":      ConvStr2Ctrlchr = vbCr
        Case "vbLf":      ConvStr2Ctrlchr = vbLf
        Case "vbNewLine": ConvStr2Ctrlchr = vbNewLine
        Case Else:        ConvStr2Ctrlchr = sValue
    End Select
End Function

' ==================================================================
' = �T�v    �f�B���N�g�����쐬����B�e�f�B���N�g����������������B
' = ����    sDirPath    String  [in]  �t�H���_�p�X
' = �ߒl    �Ȃ�
' = �o��    �t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
    
    '�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
    
    '�f�B���N�g���쐬
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
    
    Set oFileSys = Nothing
End Function

' ==================================================================
' = �T�v    �t�@�C��/�t�H���_�p�X�ꗗ���擾����(Collection,Dir�R�}���h��)
' = ����    sTrgtDir        String              [in]    �Ώۃt�H���_
' = ����    cFileList       Object(Collection)  [out]   �t�@�C��/�t�H���_�p�X�ꗗ
' = ����    lFileListType   Long                [in]    �擾����ꗗ�̌`��
' =                                                         0�F����
' =                                                         1:�t�@�C��
' =                                                         2:�t�H���_
' =                                                         ����ȊO�F�i�[���Ȃ�
' = ����    sFileExtStr     String              [in]    �擾����t�@�C���̊g���q(�ȗ��\)
' =                                                       ex1) ""
' =                                                       ex2) "*"
' =                                                       ex3) "*.c"
' =                                                       ex4) "*.txt *.log *.csv"
' = �ߒl    �Ȃ�
' = �o��    �EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' = �o��    �EsFileExtStr�̓t�@�C���w�莞�̂ݗL��
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
' ==================================================================
Public Function GetFileListCmdClct( _
    ByVal sTrgtDir As String, _
    ByRef cFileList As Object, _
    ByVal lFileListType As Long, _
    Optional ByVal sFileExtStr As String = "" _
)
    Dim objFSO As Object 'FileSystemObject�̊i�[��
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dir �R�}���h���s�i�o�͌��ʂ��ꎞ�t�@�C���Ɋi�[�j
    Dim sTmpFilePath As String
    Dim sExecCmd As String
    sTmpFilePath = CreateObject("WScript.Shell").CurrentDirectory & "\Dir.tmp"
    Dim sTrgtDirStr As String
    If sFileExtStr = "" Then
        sTrgtDirStr = """" & sTrgtDir & """"
    Else
        Dim vFileExtentions As Variant
        vFileExtentions = Split(sFileExtStr, " ")
        Dim lSplitIdx As Long
        For lSplitIdx = 0 To UBound(vFileExtentions)
            If sTrgtDirStr = "" Then
                sTrgtDirStr = """" & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            Else
                sTrgtDirStr = sTrgtDirStr & " """ & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            End If
        Next lSplitIdx
    End If
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile As Object
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(sTmpFilePath, 1)
        If Err.Number = 0 Then
            Do Until objFile.AtEndOfStream
                cFileList.Add objFile.ReadLine
            Loop
        Else
            MsgBox "�t�@�C�����J���܂���: " & Err.Description
        End If
        Set objFile = Nothing   '�I�u�W�F�N�g�̔j��
    Else
        MsgBox "�G���[ " & Err.Description
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    '�I�u�W�F�N�g�̔j��
    On Error GoTo 0
End Function
    Private Sub Test_GetFileListCmdClct()
        Dim sRootDir As String
        sRootDir = "C:\codes"
        
        Dim cFileList As Object
        Set cFileList = CreateObject("System.Collections.ArrayList")
        
'        Call GetFileListCmdClct(sRootDir, cFileList, 0)
        Call GetFileListCmdClct(sRootDir, cFileList, 1)
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*.c *.h")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*.vbs")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "")
'        Call GetFileListCmdClct(sRootDir, cFileList, 2)
        Stop
    End Sub

' ==================================================================
' = �T�v    �S�Ẵ}�N��/�v���V�[�W�����G�N�X�|�[�g����
' = ����    bTargetBook     Workbook    [in]    �G�N�X�|�[�g�Ώۃu�b�N
' = �ߒl    �Ȃ�
' = �o��    �E�ȉ��̎Q�Ɛݒ��ǉ�����K�v����B
' =           - [�c�[��] -> [�Q�Ɛݒ�] ->�uMicrosoft Visual Basic for Applications Extensibility�v
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Private Function ExportAllModules( _
    ByRef bTargetBook As Workbook _
)
    ' �t�H���_�쐬
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim sExportDirPath As String
    sExportDirPath = bTargetBook.Path & "\" & bTargetBook.Name & ".bas"
    If Not objFSO.FolderExists(sExportDirPath) Then
        objFSO.CreateFolder (sExportDirPath)
    End If
    
    Debug.Print "*** Export all macros ***"
    Debug.Print "Target book : " & bTargetBook.Name
    Debug.Print "Export path : " & bTargetBook.Path
    Dim objModule As VBComponent
    For Each objModule In bTargetBook.VBProject.VBComponents
        ' ���W���[����ʔ���
        Dim sExtension
        Select Case objModule.Type
            Case vbext_ct_ClassModule:  sExtension = "cls"
            Case vbext_ct_MSForm:       sExtension = "frm"
            Case vbext_ct_StdModule:    sExtension = "bas"
            Case vbext_ct_Document:     sExtension = "cls"
            Case Else:                  sExtension = ""
        End Select
        
        ' �G�N�X�|�[�g���{
        Dim sExportDstFilePath
        sExportDstFilePath = sExportDirPath & "\" & objModule.Name & "." & sExtension
        If sExtension = "" Then
            Debug.Print "[Ignore  ] " & objModule.Name
        Else
            Call objModule.Export(sExportDstFilePath)
            Debug.Print "[Exported] " & objModule.Name & "." & sExtension
        End If
    Next
    Debug.Print ""
End Function

' ==================================================================
' = �T�v    �w��͈͂̃Z���A�h���X(e.g. A1)�����������������
' =         �R�s�[(�N���b�v�{�[�h�Ɋi�[)����B
' =         �Ⴆ�΁AB2�`D2�͈̔͂��w�肳�ꂽ�ꍇ�A"B2&C2&D2"���R�s�[����B
' = ����    rRange          Range   [in]    �Z���͈�
' = ����    bAbsRefRow      Boolean [in]    �s�ɑ΂����ΎQ�Ǝw�� (�ȗ���)
' = ����    bAbsRefClm      Boolean [in]    ��ɑ΂����ΎQ�Ǝw�� (�ȗ���)
' = ����    bRefStyleR1C1   Boolean [in]    R1C1�`���w�� (�ȗ���)
' = ����    sDelimiter      String  [in]    ��؂蕶�� (�ȗ���)
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Macros.bas
' ==================================================================
Private Function CopyConcatedCellAddresses( _
    ByRef rRange As Range, _
    Optional ByVal bAbsRefRow As Boolean = False, _
    Optional ByVal bAbsRefClm As Boolean = False, _
    Optional ByVal bRefStyleR1C1 = False, _
    Optional ByVal sDelimiter = "" _
)
    ' �͈̓`�F�b�N
    If rRange.Columns.Count > 1 And rRange.Rows.Count > 1 Then
        MsgBox "[ERROR] 1�s�܂���1����w�肵�Ă�������"
        Return
    End If
    
    ' �Z���A�h���X�擾������
    Dim sConcatCellAdr As String
    sConcatCellAdr = ""
    Dim rCell As Range
    For Each rCell In rRange
        Dim sCellAdr As String
        Dim xlRefStyle As XlReferenceStyle
        If bRefStyleR1C1 = True Then
            xlRefStyle = xlR1C1
        Else
            xlRefStyle = xlA1
        End If
        sCellAdr = rCell.Address( _
            RowAbsolute:=bAbsRefRow, _
            ColumnAbsolute:=bAbsRefClm, _
            ReferenceStyle:=xlRefStyle _
        )
        Dim sDlmStr As String
        If sDelimiter = "" Then
            sDlmStr = "&"
        Else
            sDlmStr = "&""" & sDelimiter & """&"
        End If
        If sConcatCellAdr = "" Then
            sConcatCellAdr = sCellAdr
        Else
            sConcatCellAdr = sConcatCellAdr & sDlmStr & sCellAdr
        End If
    Next
    
    ' �N���b�v�{�[�h�ݒ�
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText sConcatCellAdr
        .PutInClipboard
    End With
    MsgBox sConcatCellAdr
End Function


