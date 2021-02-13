Option Explicit

' リポジトリへのショートカットファイル作成スクリプト v1.0

'<<使い方>>
'  1. 本スクリプトと同階層のフォルダにリポジトリパスを記載した
'     入力ファイルを作成する。
'       [ファイル名]
'         <INPUT_PATH_FILE_NAME 記載のファイル名>.txt
'       [中身の例]
'         １行目：file:///C:/svn_repo/c/FreeRTOSV7.1.1/Source
'         ２行目：file:///C:/svn_repo/c/simple_sio
'  2. 本スクリプトを実行。
'  
'  → 本スクリプトと同階層のフォルダにショートカットファイルが作成される。

Const INPUT_PATH_FILE_NAME = "_repository_path"

Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ExtractTailWord()

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sInputPathFilePath
Dim sCurDirPath
sCurDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
sInputPathFilePath = sCurDirPath & "\" & INPUT_PATH_FILE_NAME & ".txt"

If objFSO.FileExists( sInputPathFilePath ) Then
    'Do Nothing
Else
    MsgBox "「" & INPUT_PATH_FILE_NAME & ".txt」が存在しません！"
    MsgBox "処理を中断します。"
    WScript.Quit
End If

Dim cRepoFilePaths
Set cRepoFilePaths = CreateObject("System.Collections.ArrayList")

'リポジトリファイルパス取得
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile( sInputPathFilePath, 1, True ) '第二引数：IOモード（1:読出し、2:新規書込み、8:追加書込み）、第三引数：新しいファイルを作成するかどうか
Do Until objTxtFile.AtEndOfStream
    cRepoFilePaths.Add objTxtFile.ReadLine
Loop
objTxtFile.Close

'ショートカット作成
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim sRepoDirPath
For Each sRepoDirPath In cRepoFilePaths
    'MsgBox sRepoDirPath
    Dim sShortcutFilePath
    Dim sShortcutTrgtPath
    Dim sRepoDirName
    sRepoDirName = ExtractTailWord( sRepoDirPath, "/" )
    sShortcutFilePath = sCurDirPath & "\" & sRepoDirName & ".lnk"
    sShortcutTrgtPath = sRepoDirPath
    'MsgBox sRepoDirName & vbNewLine & sShortcutFilePath & vbNewLine & sShortcutTrgtPath
    
    With objWshShell.CreateShortcut( sShortcutFilePath )
        .TargetPath = "TortoiseProc.exe"
        .Arguments = " /command:repobrowser /path:""" & sShortcutTrgtPath & """"
        .Description = sShortcutTrgtPath
        .Save
    End With
Next

MsgBox "リポジトリへのショートカットを作成しました。"

' 外部プログラム インクルード関数
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

