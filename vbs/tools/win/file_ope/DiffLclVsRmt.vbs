Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" ) 'DownloadFile()

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
Dim sDownloadTrgtDirPath
Dim sScpProgramPath
Dim sDiffProgramPath
Dim sCodesDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sScpProgramPath = GetEnvVariable("MYEXEPATH_WINSCP")
'sScpProgramPath = "C:\Users\draem\Programs\program\prg_exe\WinSCP\WinSCP.exe" ��debug
sDiffProgramPath = GetEnvVariable("MYEXEPATH_WINMERGE")
sCodesDirPath = GetEnvVariable("MYDIRPATH_CODES")

'=== �t�@�C����M(Remote �� Local) ===
Dim vAnswer
vAnswer = MsgBox("�_�E�����[�h���J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbOk Then
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.vimrc "     & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.bashrc "    & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""get " & sHomeDirPath & "/.inputrc "   & sDownloadTrgtDirPath & "\"" ""exit""", 0, True
Else
    MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== �t�@�C���o�b�N�A�b�v ===
objFSO.CopyFile sDownloadTrgtDirPath & "\.vimrc",   sDownloadTrgtDirPath & "\.vimrc_rmtorg"
objFSO.CopyFile sDownloadTrgtDirPath & "\.bashrc",  sDownloadTrgtDirPath & "\.bashrc_rmtorg"
objFSO.CopyFile sDownloadTrgtDirPath & "\.inputrc", sDownloadTrgtDirPath & "\.inputrc_rmtorg"

'=== �t�H���_��r ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.inputrc"    & """ """ & sDownloadTrgtDirPath & "\.inputrc"  & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\linux\.bashrc"     & """ """ & sDownloadTrgtDirPath & "\.bashrc"   & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\.vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\_vimrc"        & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sCodesDirPath & "\vim\_gvimrc"       & """ """ & sDownloadTrgtDirPath & "\.vimrc"    & """", 10, False
vAnswer = MsgBox("��r/�}�[�W������������OK�������Ă��������B", vbOkOnly, sOutputMsg)

'=== �t�@�C�����M(Local �� Remote) ===
vAnswer = MsgBox("�ҏW�����t�@�C���𑗐M���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.vimrc" & """ ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.bashrc" & """ ""exit""", 0, True
    objWshShell.Run sScpProgramPath & " /console /command ""option batch on"" ""open " & sUserName & ":" & sPassword & "@" & sLoginServerName & """ ""cd"" ""put "& sDownloadTrgtDirPath & "\.inputrc" & """ ""exit""", 0, True
End If

'=== �t�@�C���폜 ===
vAnswer = MsgBox("�_�E�����[�h�����t�@�C�����폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.vimrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.vimrc_rmtorg"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.bashrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.bashrc_rmtorg"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.inputrc"
    objFSO.DeleteFile sDownloadTrgtDirPath & "\.inputrc_rmtorg"
End If

MsgBox "�������������܂����I", vbYesNo, sOutputMsg

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

