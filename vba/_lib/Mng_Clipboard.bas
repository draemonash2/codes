Attribute VB_Name = "Mng_Clipboard"
Option Explicit

' clipboard library v2.0

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
'GlobalALock
Private Const GHND = &H42
Private Const CF_TEXT = &H1
Private Const CF_LINK = &HBF00
Private Const CF_BITMAP = 2
Private Const CF_METAFILE = 3
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const MAXSIZE = 4096

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
    Private Function Test_SetToClipboard()
        Dim bResult As Boolean
        bResult = SetToClipboard("cliptest" & vbNewLine & "test"): Debug.Print bResult
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
    Private Function Test_GetFromClipboard()
        Dim sStr As String
        Dim bResult As Boolean
        bResult = GetFromClipboard(sStr): Debug.Print bResult & ":" & sStr
    End Function

