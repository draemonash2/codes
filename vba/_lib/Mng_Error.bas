Attribute VB_Name = "Mng_Error"
Option Explicit
 
' error ribrary v1.01
 
'************************************************************
'* �\���̒�`
'************************************************************
Public Enum E_ERROR_PROC
    ERROR_PROC_THROUGH = 0 '�G���[�o�͌���������ē��삵������
    ERROR_PROC_STOP        '�G���[�o�͌�ɒ�~����
End Enum
 
'************************************************************
'* ���W���[���� �ϐ���`
'************************************************************
Private gasErrorMsg() As String '�G���[���b�Z�[�W�i�[�̈�
 
'************************************************************
'* �֐���`
'************************************************************
' ==================================================================
' = �T�v    ����������
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Error.bas
' ==================================================================
Public Function ErrorMngInit()
    Erase gasErrorMsg
End Function
 
' ==================================================================
' = �T�v    �G���[���b�Z�[�W�i�[(�o�͂��Ȃ�)
' = ����    sErrMsg     [in]    String  �G���[���b�Z�[�W
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Error.bas
' ==================================================================
Public Function StoreErrorMsg( _
    ByVal sErrMsg As String _
)
    Dim lErrMsgNum As Long
    If Sgn(gasErrorMsg) = 0 Then
        lErrMsgNum = 0
    Else
        lErrMsgNum = UBound(gasErrorMsg) + 1
    End If
    ReDim Preserve gasErrorMsg(lErrMsgNum)
    gasErrorMsg(lErrMsgNum) = sErrMsg
End Function
 
' ==================================================================
' = �T�v    �G���[���b�Z�[�W�o��(�i�[�����G���[��S�ďo��)
' = ����    eErrorProc  [in]    E_ERROR_PROC    �o�̓G���[����
' = �ߒl    �Ȃ�
' = �ˑ�    Mng_Error.bas/ExecuteErrorProc()
' = ����    Mng_Error.bas
' ==================================================================
Public Function OutpErrorMsg( _
    ByVal eErrorProc As E_ERROR_PROC _
)
    Dim lErrMsgIdx As Long
    Dim sOutpMsg As String
 
    '�G���[�������̂ݏo��
    If Sgn(gasErrorMsg) = 0 Then
        'Do Nothing
    Else
        '#### �G���[�i�[ ####
        sOutpMsg = ""
        For lErrMsgIdx = 0 To UBound(gasErrorMsg)
            sOutpMsg = sOutpMsg & _
                                "�yErrorNo." & lErrMsgIdx + 1 & "�z" & vbCrLf & _
                                gasErrorMsg(lErrMsgIdx) & vbCrLf & vbCrLf
        Next lErrMsgIdx
 
        '#### �G���[�o�� ####
        If eErrorProc = ERROR_PROC_THROUGH Then
            MsgBox sOutpMsg, vbExclamation
        Else
            MsgBox sOutpMsg, vbCritical
        End If
        Call ErrorMngInit
 
        '#### �G���[���������� ####
        Call ExecuteErrorProc(eErrorProc)
    End If
 
End Function
 
' ==================================================================
' = �T�v    �G���[�������Ɏ��s���鏈�����Ǘ�����B
' = ����    eErrorProc  [in]    E_ERROR_PROC    �o�̓G���[����
' = �ߒl    �Ȃ�
' = �ˑ�    ��/ChkExecTerminate()
' = ����    Mng_Error.bas
' ==================================================================
Private Function ExecuteErrorProc( _
    ByVal eErrorProc As E_ERROR_PROC _
)
    Dim lProcSel As Long
    If eErrorProc = ERROR_PROC_THROUGH Then
        lProcSel = MsgBox("�������p�����܂����H", vbOKCancel)
        If lProcSel = vbOK Then
            MsgBox "�������p�����܂��I", vbExclamation
        Else
            MsgBox "�����𒆒f���܂��I", vbCritical
            Call ChkExecTerminate
            End
        End If
    Else
        MsgBox "�����𒆒f���܂��I", vbCritical
        Call ChkExecTerminate
        End
    End If
End Function

