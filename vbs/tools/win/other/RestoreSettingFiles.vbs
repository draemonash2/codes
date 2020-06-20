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
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileOrFolder()
                                                    'DeleteEmptyFolder()
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )     'ExecRunas()
Call Include( "C:\codes\vbs\_lib\Log.vbs" )         'class LogMng

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
    Call oLog.Open( _
        WScript.Arguments(ARG_IDX_LOGDIR), _
        "+w" _
    )
ElseIf WScript.Arguments.Count = ARG_COUNT_LOGINVALID Then
    'Do Nothing
Else
    oLog.Puts "-      : [error  ] argument number error! arg num is " & WScript.Arguments.Count & chr(9) & sSrcPath & chr(9) & sDstPath
    WScript.Quit
End If

If WScript.Arguments(ARG_IDX_RUNAS) = "/ExecRunas" Then
    'Do Nothing
Else
    oLog.Puts "-      : [error  ] runas exec error!"
    WScript.Quit
End If

Dim sSrcPath
Dim sDstPath
sSrcPath = WScript.Arguments(ARG_IDX_SRCPATH)
sDstPath = WScript.Arguments(ARG_IDX_DSTPATH)

Dim sFileType
Dim lRet
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_DSTPATH) )
If lRet = 2 Then
    sFileType = "folder"
ElseIf lRet = 1 Then
    sFileType = "file"
Else
    oLog.Puts "-      : [error  ] destination path is missing!        " & chr(9) & sSrcPath & chr(9) & sDstPath
    WScript.Quit
End If

'###############################################
'# �{����
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath
Dim sDstParentDirPath
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )
sShortcutPath = sDstPath & "_linksrc.lnk"

On Error Resume Next
If sFileType = "folder" Then
    If objFSO.FolderExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c rmdir /s /q """ & sSrcPath & """", 0, True
    Call ErrorCheck(1)
    If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
    Call ErrorCheck(2)
    objFSO.MoveFolder sDstPath, sSrcPath
    Call ErrorCheck(3)
    oLog.Puts "folder : [success] setting files are restored!         " & chr(9) & sSrcPath & chr(9) & sDstPath
Else
    If objFSO.FileExists( sSrcPath ) Then objWshShell.Run "%ComSpec% /c del /a /q """ & sSrcPath & """", 0, True
    Call ErrorCheck(4)
    If objFSO.FileExists( sShortcutPath ) Then objFSO.DeleteFile sShortcutPath, True
    Call ErrorCheck(5)
    objFSO.MoveFile sDstPath, sSrcPath
    Call ErrorCheck(6)
    oLog.Puts "file   : [success] setting files are restored!         " & chr(9) & sSrcPath & chr(9) & sDstPath
End If
Call DeleteEmptyFolder( sDstPath )
Call ErrorCheck(7)
On Error Goto 0

Call oLog.Close

Set oLog = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing

' = �ˑ�    �Ȃ�
' = ����    RestoreSettingFiles.vbs
Function ErrorCheck( _
    ByVal sErrorPlace _
)
    If Err.Number <> 0 Then
        oLog.Puts "-      : [error  ] an error occurred!                  " & chr(9) & sSrcPath & chr(9) & sDstPath
        oLog.Puts "           error place  : " & sErrorPlace
        oLog.Puts "           error number : " & Err.Number
        oLog.Puts "           error detail : " & Err.Description
        Err.Clear
        Call oLog.Close
        Set oLog = Nothing
        WScript.Quit
    Else
        'Do Nothing
    End If
End Function

'==========================================================
'= �C���N���[�h�֐�
'==========================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

