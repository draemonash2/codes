Option Explicit

'  [�g����]
'    SyncCodesToRemote.vbs <username> <key> <server> <home_dir_path> [<auth_type>]
'      <username>�F���[�U��
'      <key>�F�y�p�X���[�h�F�؎��z�p�X���[�h �y���J���F�؎��z�閧���i�[��(*.ppk)
'      <server>�F�ڑ���ie.g. 123.345.678.901:22�j
'      <home_dir_path>�F�z�[���f�B���N�g���p�X
'      <auth_type>�F�L�[��� (0:�p�X���[�h�F�� 1:���J���F��)�B�ȗ��B�f�t�H���g=�p�X���[�h�F�؁B

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )        'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'CreateDirectry()
                                                            'MoveToTrushBox()

'===============================================================================
'= �{����
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sUserName
Dim sKey
Dim sLoginServerName
Dim sHomeDirPath
Dim sAuthType
If WScript.Arguments.Count >= 4 Then
    sUserName = WScript.Arguments(0)
    sKey = WScript.Arguments(1)
    sLoginServerName = WScript.Arguments(2)
    sHomeDirPath = WScript.Arguments(3)
End If
If WScript.Arguments.Count >= 5 Then
    sAuthType = WScript.Arguments(4)
Else
    sAuthType = "0"
End If
If WScript.Arguments.Count < 4 Then
    MsgBox "�����𐳂����w�肵�Ă��������B", vbExclamation, WScript.ScriptName
    WScript.Quit
End If

Dim sMsgTitle
sMsgTitle = WScript.ScriptName

'=== ���O���� ===
Dim sDownloadTrgtDirPathRaw
Dim sDownloadTrgtDirPath
Dim sScpProgramPath
Dim sDiffProgramPath
Dim sCodesDirPath
sDownloadTrgtDirPathRaw = objWshShell.SpecialFolders("Desktop")
sScpProgramPath = GetEnvVariable("MYEXEPATH_WINSCP")
sDiffProgramPath = GetEnvVariable("MYEXEPATH_WINMERGE")
sCodesDirPath = GetEnvVariable("MYDIRPATH_CODES")

Dim cTrgtFileNames
Dim cTrgtDirNames
Set cTrgtFileNames = CreateObject("System.Collections.ArrayList")
Set cTrgtDirNames = CreateObject("System.Collections.ArrayList")
cTrgtDirNames.Add "vim"   : cTrgtFileNames.Add ".vimrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".bashrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".inputrc"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".gdbinit"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".tmux.conf"
cTrgtDirNames.Add "linux" : cTrgtFileNames.Add ".tigrc"

'�o�b�N�A�b�v�t�H���_�쐬
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix
Call CreateDirectry( sDownloadTrgtDirPath )

Dim vAnswer
Dim iIdx
vAnswer = MsgBox("�t�@�C������M���܂��B�_�E�����[�h���J�n���܂����H", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    On Error Resume Next '�����[�g�Ƀt�@�C�������݂��Ȃ��Ă���������
    '=== �t�@�C����M(Remote �� Local) ===
    Dim sResultMsg
    sResultMsg = ""
    For iIdx = 0 To cTrgtFileNames.Count - 1
        Dim sCmd
        If sAuthType = "0" Then ' Password
            sCmd = """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sKey & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/" & cTrgtFileNames(iIdx) & " " & sDownloadTrgtDirPath & "\"" ""exit"""
        Else                    ' PrivateKey
            sCmd = """" & sScpProgramPath & """ /console /privatekey=" & sKey & " /command ""option batch on"" ""open " & sUserName & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/" & cTrgtFileNames(iIdx) & " " & sDownloadTrgtDirPath & "\"" ""exit"""
        End If
        objWshShell.Run sCmd, 0, True
        'MsgBox sCmd
        If objFSO.FileExists(sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx)) Then
            'Do Nothing
        Else
            sResultMsg = sResultMsg & vbNewLine & cTrgtFileNames(iIdx) & " �̃_�E�����[�h�̓X�L�b�v����܂����B"
            'objFSO.CopyFile sCodesDirPath & "\" & cTrgtDirNames(iIdx) & "\" & cTrgtFileNames(iIdx), sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx)
        End If
    Next
    If sResultMsg <> "" Then
        MsgBox sResultMsg, vbOkOnly, sMsgTitle
    End If
    '=== �t�@�C���o�b�N�A�b�v ===
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objFSO.CopyFile sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx), sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & "_rmtorg"
    Next
    '=== �t�H���_��r ===
    For iIdx = 0 To cTrgtFileNames.Count - 1
        objWshShell.Run """" & sDiffProgramPath & """ -r -s """ & sCodesDirPath & "\" & cTrgtDirNames(iIdx) & "\" & cTrgtFileNames(iIdx) & """ """ & sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & """", 10, False
    Next
    On Error Goto 0
    MsgBox "��r/�}�[�W������������OK�������Ă��������B", vbOkOnly, sMsgTitle
Else
    MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sMsgTitle
    WScript.Quit
End If

vAnswer = MsgBox("�ҏW�����t�@�C���𑗐M���܂����H", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    '=== �t�@�C�����M(Local �� Remote) ===
    On Error Resume Next '���[�J���Ƀt�@�C�������݂��Ȃ��Ă���������
    For iIdx = 0 To cTrgtFileNames.Count - 1
        If sAuthType = "0" Then ' Password
            sCmd = """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sKey & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & """ ""exit"""
        Else                    ' PrivateKey
            sCmd = """" & sScpProgramPath & """ /console /privatekey=" & sKey & " /command ""option batch on"" ""open " & sUserName & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\" & cTrgtFileNames(iIdx) & """ ""exit"""
        End If
        objWshShell.Run sCmd, 0, True
        'MsgBox sCmd
    Next
    On Error Goto 0
End If

vAnswer = MsgBox("�_�E�����[�h�����t�@�C�����폜���܂����H", vbYesNo, sMsgTitle)
If vAnswer = vbYes Then
    '=== �t�H���_�폜 ===
    'objFSO.DeleteFolder sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix, True
    Call MoveToTrushBox(sDownloadTrgtDirPathRaw & "\" & "DiffLclVsRmt_" & sDateSuffix)
End If

'MsgBox "�������������܂����I", vbYesNo, sMsgTitle

'===============================================================================
'= �����֐�
'===============================================================================
' ==================================================================
' = �T�v    ���ϐ����擾����
' = ����    sEnvVar     String  [in]    ���ϐ���
' = �ߒl                String          ���ϐ��l
' = �o��    �E���ϐ������݂��Ȃ��ꍇ�A�����𒆒f����
' = �ˑ�    �Ȃ�
' = ����    �{�X�N���v�g
' ==================================================================
Private Function GetEnvVariable( _
    ByVal sEnvVar _
)
    Dim sGetValue
    sGetValue = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%" & sEnvVar & "%")
    If InStr(sGetValue, "%") > 0 then
        MsgBox "���ϐ��u" & sEnvVar & "�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sMsgTitle
        WScript.Quit
    End If
    GetEnvVariable = sGetValue
End Function

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

