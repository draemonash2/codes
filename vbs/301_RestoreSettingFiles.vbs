Option Explicit

'<<�T�v>>
'  �i�[���ύX�����v���O�����̐ݒ�t�@�C�����A���̊i�[��ɖ߂��B
'  ���̍ہA�쐬�����V���{���b�N�����N���폜���āA�i�[��ύX�O��
'  ��Ԃɕ�������B
'
'<<����>>
'  �����P�F�ޔ����t�@�C��/�t�H���_�p�X
'  �����Q�F�ޔ��t�@�C��/�t�H���_�p�X
'�i�����R�F���O�t�@�C���p�X�j
'          ���w�肵�Ȃ��ꍇ�A���O���b�Z�[�W��W���o�͂���B
'
'<<������>>
'  �P�D�V���{���b�N�����N�폜
'  �Q�D�ޔ����t�H���_�ւ̃V���[�g�J�b�g�폜
'  �R�D�t�@�C��/�t�H���_�ړ��i�ޔ��ˑޔ����j
'  �S�D�ޔ����ɍ쐬�����t�H���_���폜
'
'<<�o��>>
'  �E�ޔ��t�@�C��/�t�H���_�p�X�����݂��Ȃ��ꍇ�A�������Ȃ��B
'  �E�w�肷��p�X�̓t�@�C��/�t�H���_�ǂ���ł��B
'  �E�{�X�N���v�g���ŋ����I�ɊǗ��Ҍ����ɕύX���邽�߁A
'    ���[�J�������ł����s�ł���B�������A�{�X�N���v�g�Ăяo������
'    �Ǘ��Ҍ������s�̊m�F�E�B���h�E���\������邽�߁A�Ăяo������
'    ���炩���ߊǗ��Ҍ����Ŏ��s���Ă������Ƃ������߂���B

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\lib\Windows.vbs" )
Call Include( sMyDirPath & "\lib\Log.vbs" )

'==========================================================
'= �{����
'==========================================================
Const ARG_COUNT_LOGVALID = 4
Const ARG_COUNT_LOGINVALID = 3
Const ARG_IDX_RUNAS = 0
Const ARG_IDX_SRCPATH = 1
Const ARG_IDX_DSTPATH = 2
Const ARG_IDX_LOGDIR = 3

'�{�X�N���v�g���Ǘ��҂Ƃ��Ď��s������
If ExecRunas( False ) Then WScript.Quit

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'###############################################
'# ���O����
'###############################################
Dim oLog
Set oLog = New LogMng
If WScript.Arguments.Count = ARG_COUNT_LOGVALID Then
	Call oLog.LogFileOpen( _
		WScript.Arguments(ARG_IDX_LOGDIR), _
		"+w" _
	)
ElseIf WScript.Arguments.Count = ARG_COUNT_LOGINVALID Then
	'Do Nothing
Else
	oLog.LogPuts "#########################################################"
	oLog.LogPuts "### result : [error  ] argument number error! arg num is " & WScript.Arguments.Count
	WScript.Quit
End If

If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
	'Do Nothing
Else
	oLog.LogPuts "#########################################################"
	oLog.LogPuts "### result : [error  ] runas exec error!"
	WScript.Quit
End If

oLog.LogPuts "#########################################################"
oLog.LogPuts "### src    : " & WScript.Arguments(ARG_IDX_SRCPATH)
oLog.LogPuts "### dst    : " & WScript.Arguments(ARG_IDX_DSTPATH)

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_DSTPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	oLog.LogPuts "### result : [error  ] destination path is missing!"
	WScript.Quit
End If

'###############################################
'# �{����
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath
Dim sSrcPath
Dim sDstPath
Dim sDstParentDirPath
sSrcPath = WScript.Arguments(ARG_IDX_SRCPATH)
sDstPath = WScript.Arguments(ARG_IDX_DSTPATH)
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )
sShortcutPath = sDstPath & "_linksrc.lnk"

If sFileType = "folder" Then
	If objFSO.FolderExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c rmdir /s /q """ & sSrcPath & """", 0, True
	If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
	objFSO.MoveFolder sDstPath, sSrcPath
Else
	If objFSO.FileExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c del /a /q """ & sSrcPath & """", 0, True
	If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
	objFSO.MoveFile sDstPath, sSrcPath
End If
Call DeleteEmptyFolder( sDstPath )
oLog.LogPuts "### target : " & sFileType
oLog.LogPuts "### result : [success] setting files are restored!"

Call oLog.LogFileClose

Set oLog = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing

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

