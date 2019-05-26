Option Explicit

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )

'==========================================================
'= �ݒ�l
'==========================================================
'Const TRGT_NETWORKDRIVE_PATH = "\\RASPBERRYPI\pockethdd"
Const TRGT_NETWORKDRIVE_PATH = "\\RASPBERRYPI\LogitecHdd3T"
Const SEARCH_VOLUME_LAVEL = "PocketHdd"
Const BACKUP_PATH_SRC = "C:\Users\draem_000\AppData\Roaming\Apple Computer\MobileSync"
Const BACKUP_PATH_DST = "700_Evacuate_iTunes\MobileSync"

'==========================================================
'= �{����
'==========================================================
'�{�X�N���v�g���Ǘ��҂Ƃ��Ď��s������
Call RunasCheck

Dim sTrgtDrvPath
sTrgtDrvPath = ""
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'**** �l�b�g���[�N�h���C�u��T�� ****
On Error Resume Next
Dim lDrvLttrIdx
Dim sDrvLttr
Dim lDrvLttrAscStrt
Dim lDrvLttrAscLast
lDrvLttrAscStrt = asc("A")
lDrvLttrAscLast = asc("Z")
For lDrvLttrIdx = lDrvLttrAscStrt to lDrvLttrAscLast
    sDrvLttr = Chr(lDrvLttrIdx)
    If Err.Number = 0 Then
        If objFSO.DriveExists(sDrvLttr) Then
            Dim objDrive
            Set objDrive = objFSO.GetDrive(sDrvLttr)
            If objDrive.IsReady = True Then
                If LCase( objDrive.VolumeName ) = LCase( SEARCH_VOLUME_LAVEL ) Then
                    sTrgtDrvPath = objDrive.Path
                    Exit For
                Else
                    'Do Nothing
                End If
            Else
                'Do Nothing
            End If
        Else
            'Do Nothing
        End If
    Else
        MsgBox "{error] " & Err.Description
        WScript.Quit
    End If
Next
On Error Goto 0

'**** ���[�J���h���C�u��T�� ****
If sTrgtDrvPath = "" Then
    If objFSO.FolderExists( TRGT_NETWORKDRIVE_PATH ) Then
        sTrgtDrvPath = TRGT_NETWORKDRIVE_PATH
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'**** �V���{���b�N�����N�쐬 ****
Dim oWsh
Set oWsh = WScript.CreateObject("WScript.Shell")
If sTrgtDrvPath = "" Then
    MsgBox "�Ώۃt�H���_��������܂���ł����B"
Else
    '�V���{���b�N�����N�폜
    If objFSO.FolderExists( BACKUP_PATH_SRC ) Then
        oWsh.Run "%ComSpec% /c rmdir """ & BACKUP_PATH_SRC & """"
    Else
        'Do Nothing
    End If
    
    '�V���{���b�N�����N�폜
    oWsh.Run "%ComSpec% /c mklink /d """ & BACKUP_PATH_SRC & """ """ & _
             sTrgtDrvPath & "\" & BACKUP_PATH_DST & """"
    MsgBox "iTunes �o�b�N�A�b�v�i�[����ȉ��ɐݒ肵�܂����B" & vbNewLine & _
           "  �i�[��F" & Replace( sTrgtDrvPath, ":", "" )
End If

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
