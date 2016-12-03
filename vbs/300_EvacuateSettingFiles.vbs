Option Explicit

'<<�T�v>>
'  �v���O�����̐ݒ�t�@�C���̊i�[���ύX����B
'  �Ȃ��A���i�[�悩��V�i�[��Ɍ����ăV���{���b�N�����N���쐬���邽�߁A
'  �v���O�������ł̐ݒ�͕s�v�B
'
'<<����>>
'  �����P�F�ޔ����t�@�C��/�t�H���_�p�X
'  �����Q�F�ޔ��t�@�C��/�t�H���_�p�X
'�i�����R�F���O�t�@�C���p�X�j
'          ���w�肵�Ȃ��ꍇ�A���O���b�Z�[�W��W���o�͂���B
'
'<<������>>
'  �P�D�ޔ��t�@�C��/�t�H���_�폜
'  �Q�D�ޔ��t�@�C��/�t�H���_�쐬
'  �R�D�t�@�C��/�t�H���_�ړ��i�ޔ����ˑޔ��j
'  �S�D�ޔ����ˑޔ��ւ̃V���{���b�N�����N�쐬
'  �T�D�ޔ����t�H���_�ւ̃V���[�g�J�b�g���쐬
'
'<<�o��>>
'  �E���łɃV���{���b�N�����N���쐬����Ă���ꍇ�́A�������Ȃ��B
'  �E�ޔ����t�@�C��/�t�H���_�p�X�����݂��Ȃ��ꍇ�A�������Ȃ��B
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
Call Include( sMyDirPath & "\lib\String.vbs" )
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
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_SRCPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	oLog.LogPuts "### result : [error  ] source path is missing!"
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
Dim sSrcParentDirPath
Dim sDstParentDirPath
sSrcPath    = WScript.Arguments(ARG_IDX_SRCPATH)
sDstPath    = WScript.Arguments(ARG_IDX_DSTPATH)
sSrcParentDirPath = objFSO.GetParentFolderName( sSrcPath )
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )

If sFileType = "folder" Then
	If objFSO.GetFolder( sSrcPath ).Attributes And 1024 Then
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [error  ] setting files are already evacuated!"
	Else
		If objFSO.FolderExists( sDstPath ) Then objFSO.DeleteFolder sDstPath, True
		Call CreateDirectry( GetDirPath( sDstPath ) )
		objFSO.MoveFolder sSrcPath, sDstPath
		objWshShell.Run "%ComSpec% /c mklink /d """ & sSrcPath & """ """ & sDstPath & """", 0, True
		sShortcutPath = sDstParentDirPath & "\" & GetFileName( sSrcPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcParentDirPath
				.Save
			End With
		End If
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [success] setting files are evacuated!"
	End If
Else
	If objFSO.GetFile( sSrcPath ).Attributes And 1024 Then
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [error  ] setting files are already evacuated!"
	Else
		If objFSO.FileExists( sDstPath ) Then objFSO.DeleteFile sDstPath, True
		Call CreateDirectry( GetDirPath( sDstPath ) )
		objFSO.MoveFile sSrcPath, sDstPath
		objWshShell.Run "%ComSpec% /c mklink """ & sSrcPath & """ """ & sDstPath & """", 0, True
		sShortcutPath = sDstParentDirPath & "\" & GetFileName( sSrcPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcParentDirPath
				.Save
			End With
		End If
		oLog.LogPuts "### target : " & sFileType
		oLog.LogPuts "### result : [success] setting files are evacuated!"
	End If
End If

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
