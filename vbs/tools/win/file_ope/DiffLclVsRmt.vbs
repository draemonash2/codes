Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )        'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'CreateDirectry()

'===============================================================================
'= �{����
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 4 Then
    MsgBox "�����𐳂����w�肵�Ă��������B", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sUserName
Dim sPassword
Dim sLoginServerName
Dim sHomeDirPath
sUserName = WScript.Arguments(0)
sPassword = WScript.Arguments(1)
sLoginServerName = WScript.Arguments(2)
sHomeDirPath = WScript.Arguments(3)

Dim sOutputMsg
sOutputMsg = WScript.ScriptName

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

'�o�b�N�A�b�v�t�H���_�쐬
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = sDownloadTrgtDirPathRaw & "\" & sDateSuffix & "\" & sLoginServerName
Call CreateDirectry( sDownloadTrgtDirPath )

Dim vAnswer
vAnswer = MsgBox("�t�@�C������M���܂��B�_�E�����[�h���J�n���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== �t�@�C����M(Remote �� Local) ===
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.vimrc "     & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.bashrc "    & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.inputrc "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.gdbinit "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    
    '=== �t�@�C���o�b�N�A�b�v ===
    objFSO.CopyFile sDownloadTrgtDirPath & "\.vimrc",   sDownloadTrgtDirPath & "\.vimrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.bashrc",  sDownloadTrgtDirPath & "\.bashrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.inputrc", sDownloadTrgtDirPath & "\.inputrc_rmtorg"
    objFSO.CopyFile sDownloadTrgtDirPath & "\.gdbinit", sDownloadTrgtDirPath & "\.gdbinit_rmtorg"
    
    '=== �t�H���_��r ===
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.gdbinit"    & """ """ & sDownloadTrgtDirPath & "\.gdbinit"  & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.inputrc"    & """ """ & sDownloadTrgtDirPath & "\.inputrc"  & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.bashrc"     & """ """ & sDownloadTrgtDirPath & "\.bashrc"   & """", 10, False
    objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\.vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
    vAnswer = MsgBox("��r/�}�[�W������������OK�������Ă��������B", vbOkOnly, sOutputMsg)
Else
    MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If

vAnswer = MsgBox("�ҏW�����t�@�C���𑗐M���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== �t�@�C�����M(Local �� Remote) ===
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.vimrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.bashrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.inputrc" & """ ""exit""", 0, True
    objWshShell.Run """" & sScpProgramPath & """ /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.gdbinit" & """ ""exit""", 0, True
End If

vAnswer = MsgBox("�_�E�����[�h�����t�@�C�����폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    '=== �t�H���_�폜 ===
    objFSO.DeleteFolder sDownloadTrgtDirPathRaw & "\" & sDateSuffix, True
End If

'MsgBox "�������������܂����I", vbYesNo, sOutputMsg

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
        MsgBox "���ϐ��u" & sEnvVar & "�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
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

