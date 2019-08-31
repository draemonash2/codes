Option Explicit

'<<概要>>
'  プログラムの設定ファイルの格納先を変更する。
'  なお、元格納先から新格納先に向けてシンボリックリンクを作成するため、
'  プログラム側での設定は不要。
'
'<<引数>>
'  引数１：退避元ファイル/フォルダパス
'  引数２：退避先ファイル/フォルダパス
'（引数３：ログファイルパス）
'          ↑指定しない場合、ログメッセージを標準出力する。
'
'<<処理順>>
'  １．退避先ファイル/フォルダ削除
'  ２．退避先ファイル/フォルダ作成
'  ３．ファイル/フォルダ移動（退避元⇒退避先）
'  ４．退避元⇒退避先へのシンボリックリンク作成
'  ５．退避元フォルダへのショートカットを作成
'
'<<覚書>>
'  ・すでにシンボリックリンクが作成されている場合は、処理しない。
'  ・退避元ファイル/フォルダパスが存在しない場合、処理しない。
'  ・指定するパスはファイル/フォルダどちらでも可。
'  ・本スクリプト内で強制的に管理者権限に変更するため、
'    ローカル権限でも実行できる。ただし、本スクリプト呼び出し毎に
'    管理者権限実行の確認ウィンドウが表示されるため、呼び出し元で
'    あらかじめ管理者権限で実行しておくことをお勧めする。

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileOrFolder()
                                                    'CreateDirectry()
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )     'ExecRunas()
Call Include( "C:\codes\vbs\_lib\String.vbs" )      'GetDirPath()
Call Include( "C:\codes\vbs\_lib\Log.vbs" )         'class LogMng

'==========================================================
'= 本処理
'==========================================================
Const ARG_COUNT_LOGVALID = 4
Const ARG_COUNT_LOGINVALID = 3
Const ARG_IDX_RUNAS = 0
Const ARG_IDX_SRCPATH = 1
Const ARG_IDX_DSTPATH = 2
Const ARG_IDX_LOGDIR = 3

'本スクリプトを管理者として実行させる
If ExecRunas( False ) Then WScript.Quit

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'###############################################
'# 事前処理
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
'# 本処理
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
'= 関数定義
'==========================================================
' 外部プログラム インクルード関数
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

' = 依存    なし
' = 所属    EvacuateSettingFiles.vbs
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
