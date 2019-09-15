'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const OBJECT_SUFFIX = " - シンボリックリンク"

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "シンボリックリンク作成"

'*** 管理者として実行 ***
Call ExecRunas()

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
End If
'▼▼▼debug▼▼▼
'For Each sArg In cFilePaths
'    msgbox sArg
'Next
'▲▲▲debug▲▲▲

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "オブジェクトが選択されていません", vbYes, PROG_NAME
    MsgBox "処理を中断します", vbYes, PROG_NAME
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
        MsgBox "オブジェクトが存在しません", vbYes, PROG_NAME
        MsgBox "処理を中断します", vbYes, PROG_NAME
        WScript.Quit
    End If
    '▼▼▼debug▼▼▼
    'msgbox sCmd
    '▲▲▲debug▲▲▲
    call ExecDosCmd( sCmd )
Next

MsgBox "シンボリックリンクを作成しました", vbYes, PROG_NAME

'####################################################################
'### 関数定義
'####################################################################
' ==================================================================
' = 概要    管理者権限で実行する
' = 引数    なし
' = 戻値    なし
' = 戻値                Boolean     [out]   実行結果
' = 覚書    自動的に引数に影響を及ぼすため、要注意
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExecRunas()
    Dim oArgs
    Dim bIsRunas
    Dim sArgs
    
    bIsRunas = False
    sArgs = ""
    Set oArgs = WScript.Arguments
    
    ' フラグの取得
    If oArgs.Count > 0 Then
        If UCase(oArgs.item(0)) = "/RUNAS" Then
            bIsRunas = True
        End If
        sArgs = sArgs & " " & oArgs.item(0)
    End If
    
    Dim bIsExecutableOs
    bIsExecutableOs = false
    Dim oOsInfos
    Set oOsInfos = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
    Dim oOs
    For Each oOs in oOsInfos
        If Left(oOs.Version, 3) >= 6.0 Then
            bIsExecutableOs = True
        End If
    Next
    
    Dim oWshShell
    Set oWshShell = CreateObject("Shell.Application")
    ExecRunas = False
    If bIsRunas = False Then
        If bIsExecutableOs = True Then
            oWshShell.ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " /RUNAS " & sArgs, "", _
            "runas", _
            1
            ExecRunas = True
            Wscript.Quit
        End If
    End If
End Function

' ==================================================================
' = 概要    Dos コマンド実行
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExecDosCmd( _
    ByVal sCommand _
)
    Dim oExeResult
    Dim sStrOut
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
        sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
'   Call Test_ExecDosCmd()
    Private Sub Test_ExecDosCmd()
        Msgbox ExecDosCmd( "copy ""C:\Users\draem_000\Desktop\test.txt"" ""C:\Users\draem_000\Desktop\test2.txt""" )
        'Msgbox ExecDosCmd( "C:\codes\vbs\_lib\test.bat" )
    End Sub

' ==================================================================
' = 概要    ファイルかフォルダかを判定する
' = 引数    sChkTrgtPath    String      [in]    チェック対象フォルダ
' = 戻値                    Long                判定結果
' =                                                 1) ファイル
' =                                                 2) フォルダー
' =                                                 0) エラー（存在しないパス）
' = 覚書    FileSystemObject を使っているので、ファイル/フォルダの
' =         存在確認にも使用可能。
' = 依存    なし
' = 所属    FileSystem.vbs
' ==================================================================
Public Function GetFileOrFolder( _
    ByVal sChkTrgtPath _
)
    Dim oFileSys
    Dim bFolderExists
    Dim bFileExists
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    bFolderExists = oFileSys.FolderExists(sChkTrgtPath)
    bFileExists = oFileSys.FileExists(sChkTrgtPath)
    Set oFileSys = Nothing
    
    If bFolderExists = False And bFileExists = True Then
        GetFileOrFolder = 1 'ファイル
    ElseIf bFolderExists = True And bFileExists = False Then
        GetFileOrFolder = 2 'フォルダー
    Else
        GetFileOrFolder = 0 'エラー（存在しないパス）
    End If
End Function

