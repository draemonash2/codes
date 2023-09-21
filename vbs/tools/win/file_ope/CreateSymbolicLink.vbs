'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'【注意事項】
'   ・本スクリプトを実行するには、同階層に「CreateSymbolicLinkExec.vbs」を格納する必要がある。
'     
'     <CreateSymbolicLinkExec.vbs を設けた理由>
'       シンボリックリンクの作成は管理者権限が必須。
'       X-Finderから管理者権限で実行するには、シンボリックリンクを作成する処理を
'       別のスクリプト（CreateSymbolicLinkExec.vbs）として切り出し、そのスクリプトを
'       管理者権限として実行する必要がある。
'       なお、管理者権限で実行する場合は引数を渡せないため、引数をテキストファイルとして
'       書き出してから、呼び出す。

'####################################################################
'### 設定
'####################################################################
Const sARG_FILE_NAME = "CreateSymbolicLinkExecArg.txt" '名前は「CreateSymbolicLinkExec.vbs」内の設定値と合わせること
Const sEXEC_SCRIPT_FILE_NAME = "CreateSymbolicLinkExec.vbs"

'####################################################################
'### 事前処理
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecRunas2()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "シンボリックリンク作成（事前処理）"

'*** 対象ファイル/フォルダパス取得 ***
DIm cFilePaths
Dim sCurDirPath
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
    sCurDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    sCurDirPath = WScript.Env( "%MYDIRPATH_CODES%\vbs\tools\win\file_ope" )
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
    sCurDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
End If
'▼▼▼debug▼▼▼
'For Each sArg In cFilePaths
'    MsgBox sArg, vbYes, sPROG_NAME
'Next
'▲▲▲debug▲▲▲

'*** 対象ファイル/フォルダパス書出し ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sTrgtFilePath
sTrgtFilePath = objFSO.GetSpecialFolder(2) & "\" & sARG_FILE_NAME
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
Dim oObjPath
For Each oObjPath In cFilePaths
    'MsgBox oObjPath, vbYes, sPROG_NAME
    objTxtFile.WriteLine oObjPath
Next
objTxtFile.Close

'*** 管理者権限でスクリプト実行 ***
Dim sScriptPath
sScriptPath = sCurDirPath & "\" & sEXEC_SCRIPT_FILE_NAME
Call ExecRunas2(sScriptPath)

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

