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
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileOrFolder()
                                                    'CreateDirectry()
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )     'ExecRunas()
Call Include( "C:\codes\vbs\_lib\String.vbs" )      'GetDirPath()
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
lRet = GetFileOrFolder( WScript.Arguments(ARG_IDX_SRCPATH) )
If lRet = 2 Then
    sFileType = "folder"
ElseIf lRet = 1 Then
    sFileType = "file"
Else
    oLog.Puts "-      : [error  ] source path is missing!             " & chr(9) & sSrcPath & chr(9) & sDstPath
    WScript.Quit
End If

'###############################################
'# �{����
'###############################################
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sShortcutPath
Dim sSrcParentDirPath
Dim sDstParentDirPath
sSrcParentDirPath = objFSO.GetParentFolderName( sSrcPath )
sDstParentDirPath = objFSO.GetParentFolderName( sDstPath )
sShortcutPath = sDstParentDirPath & "\" & GetFileName( sSrcPath ) & "_linksrc.lnk"

On Error Resume Next
If sFileType = "folder" Then
    If objFSO.GetFolder( sSrcPath ).Attributes And 1024 Then
        oLog.Puts "folder : [-      ] setting files are already evacuated!" & chr(9) & sSrcPath & chr(9) & sDstPath
    Else
        If objFSO.FolderExists( sDstPath ) Then objFSO.DeleteFolder sDstPath, True
        Call ErrorCheck(1)
        Call CreateDirectry( GetDirPath( sDstPath ) )
        Call ErrorCheck(2)
        objFSO.MoveFolder sSrcPath, sDstPath
        Call ErrorCheck(3)
        objWshShell.Run "%ComSpec% /c mklink /d """ & sSrcPath & """ """ & sDstPath & """", 0, True
        Call ErrorCheck(4)
        If objFSO.FileExists( sShortcutPath ) Then
            'Do Nothing
        Else
            With objWshShell.CreateShortcut( sShortcutPath )
                .TargetPath = sSrcParentDirPath
                .Save
            End With
        End If
        Call ErrorCheck(6)
        oLog.Puts "folder : [success] setting files are evacuated!        " & chr(9) & sSrcPath & chr(9) & sDstPath
    End If
    Call ErrorCheck(7)
Else
    If objFSO.GetFile( sSrcPath ).Attributes And 1024 Then
        oLog.Puts "file   : [-      ] setting files are already evacuated!" & chr(9) & sSrcPath & chr(9) & sDstPath
    Else
        If objFSO.FileExists( sDstPath ) Then objFSO.DeleteFile sDstPath, True
        Call ErrorCheck(8)
        Call CreateDirectry( GetDirPath( sDstPath ) )
        Call ErrorCheck(9)
        objFSO.MoveFile sSrcPath, sDstPath
        Call ErrorCheck(10)
        objWshShell.Run "%ComSpec% /c mklink """ & sSrcPath & """ """ & sDstPath & """", 0, True
        Call ErrorCheck(11)
        If objFSO.FileExists( sShortcutPath ) Then
            'Do Nothing
        Else
            With objWshShell.CreateShortcut( sShortcutPath )
                .TargetPath = sSrcParentDirPath
                .Save
            End With
        End If
        Call ErrorCheck(13)
        oLog.Puts "file   : [success] setting files are evacuated!        " & chr(9) & sSrcPath & chr(9) & sDstPath
    End If
    Call ErrorCheck(14)
End If
On Error Goto 0

Call oLog.Close

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

' = �ˑ�    �Ȃ�
' = ����    EvacuateSettingFiles.vbs
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
