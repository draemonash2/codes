Attribute VB_Name = "Mng_Clipboard"
Option Explicit

' clipboard library v1.1

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

