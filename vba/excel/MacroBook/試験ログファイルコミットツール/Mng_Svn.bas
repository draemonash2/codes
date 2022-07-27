Attribute VB_Name = "Mng_Svn"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private gbIsForceExec As Boolean '�����R�~�b�g���[�h�i�G���[�L���Ɋւ�炸�A�R�~�b�g���������{������j

Public Enum E_SVN_MOD_STATUS
    MOD_STAT_OUTOFVERCTRL = 0
    MOD_STAT_NOTCHANGE
    MOD_STAT_ADDED
    MOD_STAT_MODIFIED
    MOD_STAT_OTHER
End Enum

Public Type T_SVN_MOD_STAT_INFO
    sPath As String
    eSvnModStat As E_SVN_MOD_STATUS
End Type

Public Function SvnInit()
    gbIsForceExec = False
End Function

Public Function ExecCommit()
    Dim bIsCommitExec As Boolean

    '=== �R�~�b�g�O�`�F�b�N����
    bIsCommitExec = PreChkBeforeCommit()
    
    '=== �R�~�b�g���� ===
    If bIsCommitExec = True Then
        Call AddFiles2Svn
        Call ShowCommitDialog
    Else
        Call OutpErrorMsg(ERROR_PROC_STOP) '�G���[�o�͌�A������~
    End If
End Function

Private Function PreChkBeforeCommit() As Boolean
    Dim bIsCommitExec As Boolean
    Dim lProcSel As Long
    Dim bChkProcFin As Boolean
    
    bIsCommitExec = True
    bChkProcFin = False
    
    '=== �G���[�����m�F ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If gbIsForceExec = True Then '�����R�~�b�g���[�h�� ON �̎��̓G���[�𖳎�����I
            bChkProcFin = True
            bIsCommitExec = True
        Else
            If ErrorExist() = False Then
                'Do Nothing
            Else
                Call StoreErrorMsg( _
                    "�G���[���������Ă��܂��I" & vbNewLine & _
                    "���ڏ��������̓��O�t�@�C�����m�F���Ă��������I" _
                )
                bChkProcFin = True
                bIsCommitExec = False
            End If
        End If
    End If
    
    '=== ���[�j���O�����m�F ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If gbIsForceExec = True Then '�����R�~�b�g���[�h�� ON �̎��̓��[�j���O�𖳎�����I
            bChkProcFin = True
            bIsCommitExec = True
        Else
            If WarningExist() = False Then
                'Do Nothing
            Else
                lProcSel = MsgBox( _
                                    "���[�j���O���������Ă��܂����A�����𑱂��܂����H" & vbNewLine & _
                                    "" & vbNewLine & _
                                    "���[�j���O�̓��e���m�F�������ꍇ�́u�L�����Z���v���A" & vbNewLine & _
                                    "���m�̖��ł���΁uOK�v�������Ă��������B", _
                                    vbOKCancel _
                           )
                If lProcSel = vbOK Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg("�������L�����Z������܂����I")
                    bChkProcFin = True
                    bIsCommitExec = False
                End If
            End If
        End If
    End If
    
    '=== CUI �� SVN �C���X�g�[���m�F ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If ChkCuiSvnInstall() = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                "CUI �� Subversion ���C���X�g�[������Ă���܂���I" & vbNewLine & _
                "�C���X�g�[����ɍēx���s���Ă��������I" _
            )
            bChkProcFin = True
            bIsCommitExec = False
        End If
    End If
    
    '=== �ŏI�m�F ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        MsgBox "���O�t�H���_�\���E���O�t�@�C�����ɖ�肪�Ȃ��������߁A" & vbNewLine & _
               "���O�t�@�C���� SVN ADD ���ăR�~�b�g�_�C�A���O��\�����܂��B"
        lProcSel = MsgBox( _
                            "�������O�t�@�C���� SVN ADD ���܂��B" & vbNewLine & _
                            "��肪�Ȃ���΁uOK�v�������Ă��������B" & vbNewLine & _
                            "" & vbNewLine & _
                            "�iADD �ς݂̏ꍇ���uOK�v�Ƃ��Ă��������j", _
                            vbOKCancel _
                   )
        If lProcSel = vbOK Then
            'Do Nothing
        Else
            Call StoreErrorMsg("�������L�����Z������܂����I")
            bChkProcFin = True
            bIsCommitExec = False
        End If
    End If
    
    PreChkBeforeCommit = bIsCommitExec
End Function

Private Function ChkCuiSvnInstall() As Boolean
    Dim sCmdRslt As String
    sCmdRslt = ExecDosCmd("svn --version")
    'Debug.Print sCmdRslt
    If InStr(sCmdRslt, "svn, version") > 0 And _
       InStr(sCmdRslt, "Copyright") > 0 Then
        ChkCuiSvnInstall = True
    ElseIf sCmdRslt = "" Then
        ChkCuiSvnInstall = False
    Else
        Stop
    End If
End Function

Private Function AddFiles2Svn()
    Dim sCmdRslt As String
    Dim sCmd As String
    
    '*** ���O�t�H���_�z�� ***
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "svn add --parents " & gtInputInfo.sTestLogDirPath
    Else
        sCmd = "svn add --parents " & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName
    End If
    sCmdRslt = ExecDosCmd(sCmd)
    '���ɒǉ��ς݂̏ꍇ�AsCmdRslt �� "" �ƂȂ��Ă��܂����߁A
    '�G���[�`�F�b�N���Ȃ��B
'    If sCmdRslt = "" Then
'        Call StoreErrorMsg( _
'            "svn add �R�}���h�����s���܂����I" _
'        )
'        Call OutpErrorMsg(ERROR_PROC_STOP)
'    Else
'        'Do Nothing
'    End If
    
    '*** �������ڏ� ***
    sCmd = "svn add --parents " & gtInputInfo.sTestDocFilePath
    sCmdRslt = ExecDosCmd(sCmd)
    '���ɒǉ��ς݂̏ꍇ�AsCmdRslt �� "" �ƂȂ��Ă��܂����߁A
    '�G���[�`�F�b�N���Ȃ��B
'    If sCmdRslt = "" Then
'        Call StoreErrorMsg( _
'            "svn add �R�}���h�����s���܂����I" _
'        )
'        Call OutpErrorMsg(ERROR_PROC_STOP)
'    Else
'        'Do Nothing
'    End If
End Function

Private Function ShowCommitDialog()
    Dim sCmdRslt As String
    Dim sCmd As String
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "��" & """ "
    Else
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "��" & """ "
    End If
    sCmdRslt = ExecDosCmd(sCmd)
End Function

Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    'Debug.Print sStrOut
    'Debug.Print ""
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function

Public Function GetSvnModStatList( _
    ByVal sFilePath As String _
) As T_SVN_MOD_STAT_INFO()
    Dim sCmd As String
    Dim sCmdRslt As String
    Dim asCmdRslt() As String
    Dim lCmdRsltCnt As Long
    Dim sExtFilePath As String
    Dim sModStatus As String
    Dim eModStatus As E_SVN_MOD_STATUS
    Dim sCmdRsltLine As String
    Dim atSvnModStat() As T_SVN_MOD_STAT_INFO
    Dim lSvnModStatIdx As Long
    Const E_PATH_STRT_POS = 42
    
    sCmd = "svn status -v " & sFilePath
    sCmdRslt = ExecDosCmd(sCmd)
    asCmdRslt = Split(sCmdRslt, vbNewLine)
    For lCmdRsltCnt = 0 To UBound(asCmdRslt)
        sCmdRsltLine = asCmdRslt(lCmdRsltCnt)
        If sCmdRsltLine = "" Then
            'Do Nothing
        Else
            If Sgn(atSvnModStat) = 0 Then
                lSvnModStatIdx = 0
            Else
                lSvnModStatIdx = lSvnModStatIdx + 1
            End If
            ReDim Preserve atSvnModStat(lSvnModStatIdx)
            sModStatus = Mid$(sCmdRsltLine, 1, 1)
            sExtFilePath = Mid$(sCmdRsltLine, E_PATH_STRT_POS, Len(sCmdRsltLine) - E_PATH_STRT_POS + 1)
            Select Case sModStatus
                Case " ": eModStatus = MOD_STAT_NOTCHANGE
                Case "?": eModStatus = MOD_STAT_OUTOFVERCTRL
                Case "A": eModStatus = MOD_STAT_ADDED
                Case "M": eModStatus = MOD_STAT_MODIFIED
                Case Else: eModStatus = MOD_STAT_OTHER
            End Select
            atSvnModStat(lSvnModStatIdx).sPath = sExtFilePath
            atSvnModStat(lSvnModStatIdx).eSvnModStat = eModStatus
            'Debug.Print oSvnModStatList.Item(sExtFilePath) & ":" & sExtFilePath
        End If
    Next lCmdRsltCnt
    GetSvnModStatList = atSvnModStat
End Function

'���e�X�g�p��
Sub test3()
    Dim sFilePath As String
    Dim lLoopCnt As Long
    Dim oStopWatch As New StopWatch
    Dim atSvnModStat() As T_SVN_MOD_STAT_INFO
    
    sFilePath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\Busmaster"
    
    Call SvnInit
    Call oStopWatch.StartT
    atSvnModStat = GetSvnModStatList(sFilePath)
    Debug.Print oStopWatch.StopT
End Sub

Sub test()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\commit_test\ST"
    gtInputInfo.sTestDocFilePath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\test\TF2���v���g2nd�V�X�e���������ڏ�-17650_��.xlsx"
    Call ShowCommitDialog
End Sub

Sub test2()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\commit_test\ST"
    'gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\commit_test"
    gtInputInfo.sTestDocFilePath = "C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\test\TF2���v���g2nd�V�X�e���������ڏ�-17650_��.xlsx"
    Call AddFiles2Svn
End Sub

