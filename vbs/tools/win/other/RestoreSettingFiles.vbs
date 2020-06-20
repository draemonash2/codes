Option Explicit

'<<概要>>
'  格納先を変更したプログラムの設定ファイルを、元の格納先に戻す。
'  その際、作成したシンボリックリンクを削除して、格納先変更前の
'  状態に復元する。
'
'<<引数>>
'  引数１：退避元ファイル/フォルダパス
'  引数２：退避先ファイル/フォルダパス
'（引数３：ログファイルパス）
'          ↑指定しない場合、ログメッセージを標準出力する。
'
'<<処理順>>
'  １．シンボリックリンク削除
'  ２．退避元フォルダへのショートカット削除
'  ３．ファイル/フォルダ移動（退避先⇒退避元）
'  ４．退避時に作成したフォルダを削除
'
'<<覚書>>
'  ・退避先ファイル/フォルダパスが存在しない場合、処理しない。
'  ・指定するパスはファイル/フォルダどちらでも可。
'  ・本スクリプト内で強制的に管理者権限に変更するため、
'    ローカル権限でも実行できる。ただし、本スクリプト呼び出し毎に
'    管理者権限実行の確認ウィンドウが表示されるため、呼び出し元で
'    あらかじめ管理者権限で実行しておくことをお勧めする。

'==========================================================
'= インクルード
'==========================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileOrFolder()
                                                    'DeleteEmptyFolder()
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )     'ExecRunas()
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
'# 本処理
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

' = 依存    なし
' = 所属    RestoreSettingFiles.vbs
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
'= インクルード関数
'==========================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

