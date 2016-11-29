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
Dim bIsLogValid
If WScript.Arguments.Count = ARG_COUNT_LOGVALID Then
	bIsLogValid = True
	Dim objLogFile
	Set objLogFile = objFSO.OpenTextFile( WScript.Arguments(ARG_IDX_LOGDIR), 8, True) '�������FIO���[�h�i1:�Ǐo���A2:�V�K�����݁A8:�ǉ������݁j
ElseIf WScript.Arguments.Count = ARG_COUNT_LOGINVALID Then
	bIsLogValid = False
Else
	WScript.Echo "[error] argument number error!" & vbNewLine & _
		   "  argument num : " & WScript.Arguments.Count
	WScript.Quit
End If

Dim sExecResult
If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
	'Do Nothing
Else
	sExecResult = "[error] runas exec error!"
	If bIsLogValid = True Then
		objLogFile.WriteLine sExecResult
	Else
		WScript.Echo sExecResult
	End If
	WScript.Quit
End If

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_SRCPATH) )
If lRet = 2 Then
	sFileType = "folder"
ElseIf lRet = 1 Then
	sFileType = "file"
Else
	sExecResult = "[error] source path is missing!" & vbNewLine & _
				  "  src : " & WScript.Arguments(ARG_IDX_SRCPATH) & vbNewLine & _
				  "  dst : " & WScript.Arguments(ARG_IDX_DSTPATH)
	If bIsLogValid = True Then
		objLogFile.WriteLine sExecResult
	Else
		WScript.Echo sExecResult
	End If
	WScript.Quit
End If

'###############################################
'# �{����
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath

If sFileType = "folder" Then
	Dim sSrcDirPath
	Dim sDstDirPath
	Dim sSrcDirParentDirPath
	Dim sDstDirParentDirPath
	sSrcDirPath    = WScript.Arguments(ARG_IDX_SRCPATH)
	sDstDirPath    = WScript.Arguments(ARG_IDX_DSTPATH)
	sSrcDirParentDirPath = objFSO.GetParentFolderName( sSrcDirPath )
	sDstDirParentDirPath = objFSO.GetParentFolderName( sDstDirPath )
	If objFSO.GetFolder( sSrcDirPath ).Attributes And 1024 Then
		sExecResult = "[error] setting files are already evacuated!" & vbNewLine & _
					  "  src : " & sSrcDirPath & vbNewLine & _
					  "  dst : " & sDstDirPath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	Else
		If objFSO.FolderExists( sDstDirPath ) Then objFSO.DeleteFolder sDstDirPath, True
		Call CreateDirectry( GetDirPath( sDstDirPath ) )
		objFSO.MoveFolder sSrcDirPath, sDstDirPath
		objWshShell.Run "%ComSpec% /c mklink /d """ & sSrcDirPath & """ """ & sDstDirPath & """", 0, True
		sShortcutPath = sDstDirParentDirPath & "\" & GetFileName( sSrcDirPath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcDirParentDirPath
				.Save
			End With
		End If
		sExecResult = "[success] setting files are evacuated!" & vbNewLine & _
					  "  src : " & sSrcDirPath & vbNewLine & _
					  "  dst : " & sDstDirPath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	End If
Else
	Dim sSrcFilePath
	Dim sDstFilePath
	Dim sDstFileParentDirPath
	Dim sSrcFileParentDirPath
	
	sSrcFilePath	= WScript.Arguments(ARG_IDX_SRCPATH)
	sDstFilePath	= WScript.Arguments(ARG_IDX_DSTPATH)
	sDstFileParentDirPath = objFSO.GetParentFolderName( sDstFilePath )
	sSrcFileParentDirPath = objFSO.GetParentFolderName( sSrcFilePath )
	
	If objFSO.GetFile( sSrcFilePath ).Attributes And 1024 Then
		sExecResult = "[error] setting files are already evacuated!" & vbNewLine & _
					  "  src : " & sSrcFilePath & vbNewLine & _
					  "  dst : " & sDstFilePath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	Else
		If objFSO.FileExists( sDstFilePath ) Then objFSO.DeleteFile sDstFilePath, True
		Call CreateDirectry( GetDirPath( sDstFilePath ) )
		objFSO.MoveFile sSrcFilePath, sDstFilePath
		objWshShell.Run "%ComSpec% /c mklink """ & sSrcFilePath & """ """ & sDstFilePath & """", 0, True
		sShortcutPath = sDstFileParentDirPath & "\" & GetFileName( sSrcFilePath ) & "_linksrc.lnk"
		If objFSO.FileExists( sShortcutPath ) Then
			'Do Nothing
		Else
			With objWshShell.CreateShortcut( sShortcutPath )
				.TargetPath = sSrcFileParentDirPath
				.Save
			End With
		End If
		sExecResult = "[success] setting files are evacuated!" & vbNewLine & _
					  "  src : " & sSrcFilePath & vbNewLine & _
					  "  dst : " & sDstFilePath
		If bIsLogValid = True Then
			objLogFile.WriteLine sExecResult
		Else
			WScript.Echo sExecResult
		End If
	End If
End If

If bIsLogValid = True Then
	objLogFile.Close
	Set objLogFile = Nothing
Else
	'Do Nothing
End If

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

