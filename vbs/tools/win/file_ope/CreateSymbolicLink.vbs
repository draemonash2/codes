'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'【注意事項】X-Finderから実行する場合は、管理者権限にて起動したX-Finderから起動すること

'####################################################################
'### 設定
'####################################################################
Const OBJECT_SUFFIX = " - シンボリックリンク"

'####################################################################
'### 事前処理
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecRunas()
                                                            'ExecDosCmd()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileOrFolder()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "シンボリックリンク作成"

'*** ファイル/フォルダ名取得 ***
DIm cFilePaths
If EXECUTION_MODE = 0 Then 'Explorerから実行
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    Dim sArg
    For Each sArg In WScript.Arguments
        If sArg = "/RUNAS" Then
            'Do Nothing
        Else
            cFilePaths.add sArg
        End If
    Next
    Call ExecRunas() '管理者として実行
ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
Else 'デバッグ実行
    MsgBox "デバッグモードです。"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim sDesktop
    sDesktop = objWshShell.SpecialFolders("Desktop")
    objWshShell.Run "cmd /c echo.> """ & sDesktop & "\test.txt""", 0, True
    objWshShell.Run "cmd /c mkdir """ & sDesktop & "\test2""", 0, True
    cFilePaths.Add sDesktop & "\test.txt"
    cFilePaths.Add sDesktop & "\test2"
    Call ExecRunas() '管理者として実行
End If
'▼▼▼debug▼▼▼
'For Each sArg In cFilePaths
'    msgbox sArg
'Next
'▲▲▲debug▲▲▲

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "オブジェクトが選択されていません", vbYes, sPROG_NAME
    MsgBox "処理を中断します", vbYes, sPROG_NAME
    WScript.Quit
End If

'*** シンボリックリンク作成 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim oObjPath
Dim lObjType '0:notexists 1:file 2:folder
For Each oObjPath In cFilePaths
    lObjType = GetFileOrFolder( oObjPath )
    
    Dim sDstPath
    Dim sSrcPath
    sDstPath = oObjPath
    Dim sCmd
    If lObjType = 1 Then 'file
        sSrcPath = objFSO.GetParentFolderName( oObjPath ) & "\" & _
                   objFSO.GetBaseName( oObjPath ) & OBJECT_SUFFIX & "." & _
                   objFSO.GetExtensionName( oObjPath )
        sCmd = "mklink """ & sSrcPath & """ """ & sDstPath & """"
    ElseIf lObjType = 2 Then 'folder
        sSrcPath = oObjPath & OBJECT_SUFFIX
        sCmd = "mklink /d """ & sSrcPath & """ """ & sDstPath & """"
    Else 'not exists
        MsgBox "オブジェクトが存在しません", vbYes, sPROG_NAME
        MsgBox "処理を中断します", vbYes, sPROG_NAME
        WScript.Quit
    End If
    '▼▼▼debug▼▼▼
    'msgbox sCmd
    '▲▲▲debug▲▲▲
    call ExecDosCmd( sCmd )
Next

MsgBox "シンボリックリンクを作成しました", vbYes, sPROG_NAME

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

