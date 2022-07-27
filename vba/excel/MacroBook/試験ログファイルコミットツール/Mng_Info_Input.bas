Attribute VB_Name = "Mng_Info_Input"
Option Explicit

Private Const INPUT_SRCH_KEY_TRGT_PHASE = "�����t�F�[�Y"
Private Const INPUT_SRCH_KEY_SBJCT_NAME = "�Č���"
Private Const INPUT_SRCH_KEY_DOC_PATH = "�������ڏ��t�@�C���p�X"
Private Const INPUT_SRCH_KEY_LOG_PATH = "�������O�t�H���_�p�X"
Private Const INPUT_SRCH_KEY_TESTER = "�]����"
Private Const INPUT_SRCH_KEY_TEST_DATE = "�N����"
Private Const INPUT_SRCH_KEY_TEST_RSLT = "���ʔ���"
Private Const INPUT_SRCH_KEY_TEST_DATA = "�����f�[�^"
Private Const INPUT_SRCH_KEY_REV_SRC = "���� Rev�i�\�[�X�R�[�h�j"
Private Const INPUT_SRCH_KEY_REV_HEXABS = "���� Rev�iHEX/ABS�j"
Private Const INPUT_SRCH_KEY_REV_A2L = "���� Rev�iA2L�j"

Private Const INPUT_SHEET_NAME = "�f�[�^����"

Private Const TRGT_PHASE_NAME_UT = "�P�̎���"
Private Const TRGT_PHASE_NAME_CT = "��������"
Private Const TRGT_PHASE_NAME_FT = "�@�\����"
Private Const TRGT_PHASE_NAME_ST = "�V�X�e������"

Public Enum E_TRGT_PHASE
    TRGT_PHASE_UT
    TRGT_PHASE_CT
    TRGT_PHASE_FT
    TRGT_PHASE_ST
End Enum

Public Type T_INPUT_INFO
    eTrgtPhase As E_TRGT_PHASE
    sSubjectName As String
    sTestDocFilePath As String
    sTestLogDirPath As String
    sTester As String
    sTestDate As String
    sTestRslt As String
    sRevSrc As String
    sRevHexAbs As String
    sRevA2L As String
End Type

Public gtInputInfo As T_INPUT_INFO

Public Function InputInfoInit()
    Dim tInputInfoInit As T_INPUT_INFO
    gtInputInfo = tInputInfoInit
End Function

Public Function GetInputInfo()
    Dim shTrgtSht As Worksheet
    Dim sFileBaseName As String
    Dim sTrgtPhaseName As String
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Set shTrgtSht = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    
    '### �Z���f�[�^�擾 ###
    '*** �����t�F�[�Y ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TRGT_PHASE, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_DOC_PATH & "�v���L�ڂ���Ă���܂���I")
        Call OutpErrorMsg(ERROR_PROC_STOP)
    Else
        'Do Nothing
    End If
    sTrgtPhaseName = tNearCellData.sCellValue
    Select Case sTrgtPhaseName
        Case TRGT_PHASE_NAME_UT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_UT
        Case TRGT_PHASE_NAME_CT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_CT
        Case TRGT_PHASE_NAME_FT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_FT
        Case TRGT_PHASE_NAME_ST:    gtInputInfo.eTrgtPhase = TRGT_PHASE_ST
        Case Else:                  Stop
    End Select
    
    '*** �Č��� ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_SBJCT_NAME, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_SBJCT_NAME & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sSubjectName = tNearCellData.sCellValue
    
    '*** �������ڏ��t�@�C���p�X ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_DOC_PATH, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_DOC_PATH & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestDocFilePath = tNearCellData.sCellValue
    
    '*** �������O�t�H���_�p�X ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_LOG_PATH, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_LOG_PATH & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestLogDirPath = tNearCellData.sCellValue
    
    '*** �]���� ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TESTER, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_TESTER & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTester = tNearCellData.sCellValue
    
    '*** �N���� ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TEST_DATE, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_TEST_DATE & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestDate = tNearCellData.sCellValue
    
    '*** ���ʔ��� ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TEST_RSLT, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_TEST_RSLT & "�v���L�ڂ���Ă���܂���I")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestRslt = tNearCellData.sCellValue
    
    '*** �����f�[�^ ***
    '�����f�[�^�͊i�[���Ȃ�
    
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
        '*** ���� Rev�i�\�[�X�R�[�h�j ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_SRC, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_REV_SRC & "�v���L�ڂ���Ă���܂���I")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevSrc = tNearCellData.sCellValue
    Else
        '*** ���� Rev�iHEX/ABS�j ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_HEXABS, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_REV_HEXABS & "�v���L�ڂ���Ă���܂���I")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevHexAbs = tNearCellData.sCellValue
        
        '*** ���� Rev�iA2L�j ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_A2L, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("�u" & INPUT_SRCH_KEY_REV_A2L & "�v���L�ڂ���Ă���܂���I")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevA2L = tNearCellData.sCellValue
    End If
    
    Call OutpErrorMsg(ERROR_PROC_STOP)
    
    '### �f�[�^�`�F�b�N���� ###
    '�������ڏ����̎����t�F�[�Y��v�`�F�b�N
    sFileBaseName = GetFileNameBase(gtInputInfo.sTestDocFilePath)
    If InStr(sFileBaseName, sTrgtPhaseName) > 0 Then
        'Do Nothing
    Else
        Call StoreErrorMsg("���ڏ����Ɂu" & sTrgtPhaseName & "�v���܂܂�Ă���܂���I")
        Call OutpErrorMsg(ERROR_PROC_THROUGH)
    End If
End Function

Public Function OutputDocFilePathCell( _
    ByVal sPath As String _
)
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    
    With ThisWorkbook
        '�������ڏ��t�@�C���p�X
        tNearCellData = GetNearCellData(.Sheets(INPUT_SHEET_NAME), INPUT_SRCH_KEY_DOC_PATH, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        
        .Sheets(INPUT_SHEET_NAME).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = sPath
    End With
End Function

Public Function OutputLogDirPathCell( _
    ByVal sPath As String _
)
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    
    With ThisWorkbook
        '�������ڏ��t�@�C���p�X
        tNearCellData = GetNearCellData(.Sheets(INPUT_SHEET_NAME), INPUT_SRCH_KEY_LOG_PATH, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        
        .Sheets(INPUT_SHEET_NAME).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = sPath
    End With
End Function



