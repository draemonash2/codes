'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�B���t�@�C���\���؂�ւ�"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If EXECUTION_MODE = 1 Then 'X-Finder������s
        'Do Nothing
    Else
        MsgBox "���̃X�N���v�g��X-Finder�ȊO�ł͎��s�ł��܂���B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

If bIsContinue = True Then
    If InStr( WScript.Env("Style"), "h" ) > 0 Then
        MsgBox "�B���t�@�C���A�V�X�e���t�@�C�����y��\���z�ɂ��܂��B", vbOKOnly, PROG_NAME
    Else
        MsgBox _
            "�B���t�@�C���A�V�X�e���t�@�C�����y�\���z���܂��B" & vbNewLine & _
            "" & vbNewLine & _
            "(��) �G�N�X�v���[���[�̃t�H���_�ݒ�ɂāu�ی삳�ꂽ�I�y���[�e�B���O�V�X�e���t�@�C����\�����Ȃ��i�����j�v���`�F�b�N����Ă���ꍇ�A�V�X�e���t�@�C���͕\������܂���B" _
            , vbOKOnly, PROG_NAME
    End If
    WScript.Exec("Change:Style ~h")
    WScript.Exec("Refresh:4")
Else
    'Do Nothing
End If