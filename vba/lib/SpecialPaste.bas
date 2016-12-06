Attribute VB_Name = "SpecialPaste"
Option Explicit

Const EXEC_SEND_KEY As Boolean = False
Const SENDKEY_SLEEPTIME As Long = 50

Public Sub EnableSpetialKeyMode()
    MsgBox "�ȉ��̃V���[�g�J�b�g�L�[���u�����t�������Ή����[�h�v�ɐ؂�ւ��܂��B" & vbNewLine & _
           "�EShift + Ctrl + ""+""" & vbNewLine & _
           "�ECtrl + v" & vbNewLine & _
           "�ECtrl + ""-""" & vbNewLine & _
           "" & vbNewLine & _
           "�����Ӂ� ���̃��[�h�ł̓A���h�D�ł��܂���I"
    Application.OnKey "+^{+}", "NewInsert"
    Application.OnKey "^v", "NewPaste"
    Application.OnKey "^-", "NewDelete"
End Sub
Public Sub DisableSpetialKeyMode()
    MsgBox "�ȉ��̃V���[�g�J�b�g�L�[���u�m�[�}�����[�h�v�ɐ؂�ւ��܂��B" & vbNewLine & _
           "�EShift + Ctrl + ""+""" & vbNewLine & _
           "�ECtrl + v" & vbNewLine & _
           "�ECtrl + ""-"""
    Application.OnKey "+^{+}"
    Application.OnKey "^v"
    Application.OnKey "^-"
End Sub
Private Sub NewInsert()
    '�}���\��t�������𖳌��ɂ���
    Select Case Application.CutCopyMode
        Case xlCopy
            MsgBox "�}���\��t���͖����ł��B"
        Case xlCut
            MsgBox "�}���\��t���͖����ł��B"
        Case Else
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hie", SENDKEY_SLEEPTIME) '�}��
            Else
                Application.ScreenUpdating = False
                Selection.Insert
                Application.ScreenUpdating = True
            End If
    End Select
End Sub
Private Sub NewPaste()
    Select Case Application.CutCopyMode
        Case xlCopy
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hvf", SENDKEY_SLEEPTIME) '�����\��t��
            Else
                Application.ScreenUpdating = False
                '������\��t����
                Selection.PasteSpecial _
                    Paste:=xlPasteFormulas, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
    '            '�����t���������������ē\��t����
    '            Selection.PasteSpecial _
    '                Paste:=xlPasteAllMergingConditionalFormats, _
    '                Operation:=xlNone, _
    '                SkipBlanks:=False, _
    '                Transpose:=False
                Application.ScreenUpdating = True
            End If
        Case xlCut
            MsgBox "�J�b�g���y�[�X�g�͖����ł��B"
        Case Else
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hvt", SENDKEY_SLEEPTIME) '�\��t��
            Else
                Application.ScreenUpdating = False
                Dim doDataObj As New DataObject
                doDataObj.GetFromClipboard
                Selection(1).Value = doDataObj.GetText
                Application.ScreenUpdating = True
            End If
    End Select
End Sub
Private Sub NewDelete()
    '�u��s�폜�v���́u�s�}������s�폜�v�Ƃ���B
    '�i��s�݂̂̍폜�͏����t�����������B����Ă��܂����߁j
    If Selection.Rows.Count = 1 And _
       Selection.Columns.Count = Columns.Count Then
        MsgBox "�����t������������邽�߁A�P�s�̍폜�͂ł��܂���B"
'        If EXEC_SEND_KEY = True Then
'            Call SendKeysBetweenWait("%hie", SENDKEY_SLEEPTIME) '�}��
'            Call SendKeysBetweenWait("+{DOWN}", SENDKEY_SLEEPTIME) '�V�t�g+��
'            Call SendKeysBetweenWait("%hdd", SENDKEY_SLEEPTIME) '�폜
'        Else
'            Application.ScreenUpdating = False
'            Selection.Insert
'            Selection.Resize(Selection.Rows.Count + 1).Select '�s�������g��
'            Selection.Delete
'            Selection.Resize(1).Select
'            Application.ScreenUpdating = True
'        End If
    Else
        If EXEC_SEND_KEY = True Then
            Call SendKeysBetweenWait("%hdd", SENDKEY_SLEEPTIME) '�폜
        Else
            Application.ScreenUpdating = False
            Selection.Delete
            Application.ScreenUpdating = True
        End If
    End If
End Sub
