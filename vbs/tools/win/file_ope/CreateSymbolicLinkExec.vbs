'Option Explicit

'【注意事項】
'   ・本スクリプトを実行するには、同階層に「CreateSymbolicLink.vbs」を格納する必要がある。
'     
'     <本スクリプト CreateSymbolicLinkExec.vbs を設けた理由>
'       シンボリックリンクの作成は管理者権限が必須。
'       X-Finderから管理者権限で実行するには、シンボリックリンクを作成する処理を
'       別のスクリプト（CreateSymbolicLinkExec.vbs）として切り出し、そのスクリプトを
'       管理者権限として実行する必要がある。
'       なお、管理者権限で実行する場合は引数を渡せないため、引数をテキストファイルとして
'       書き出してから、呼び出す。

'####################################################################
'### 設定
'####################################################################
Const OBJECT_SUFFIX = ".symlink"
Const sARG_FILE_NAME = "CreateSymbolicLinkExecArg.txt" '名前は「CreateSymbolicLink.vbs」内の設定値と合わせること

'####################################################################
'### 事前処理
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecDosCmd()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileOrFolder()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "シンボリックリンク作成（実処理）"

'*** 対象ファイル/フォルダ名取得 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sTrgtFilePath
sTrgtFilePath = objFSO.GetSpecialFolder(2) & "\" & sARG_FILE_NAME
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
DIm cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Do Until objTxtFile.AtEndOfStream
    cFilePaths.Add objTxtFile.ReadLine
Loop
objTxtFile.Close
objFSO.DeleteFile sTrgtFilePath, True
'▼▼▼debug▼▼▼
'Dim sArg
'For Each sArg In cFilePaths
'    MsgBox sArg, vbYes, sPROG_NAME
'Next
'WScript.Quit
'▲▲▲debug▲▲▲

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "オブジェクトが選択されていません", vbYes, sPROG_NAME
    MsgBox "処理を中断します", vbYes, sPROG_NAME
    WScript.Quit
End If

'*** シンボリックリンク作成 ***
Dim oObjPath
Dim lObjType '0:notexists 1:file 2:folder
For Each oObjPath In cFilePaths
    lObjType = GetFileOrFolder( oObjPath )
    
    Dim sDstPath
    Dim sSrcPath
    sDstPath = oObjPath
    Dim sCmd
    If lObjType = 1 Then 'file
       'sSrcPath = objFSO.GetParentFolderName( oObjPath ) & "\" & _
       '           objFSO.GetBaseName( oObjPath ) & OBJECT_SUFFIX & "." & _
       '           objFSO.GetExtensionName( oObjPath )
        sSrcPath = oObjPath & OBJECT_SUFFIX & "." & objFSO.GetExtensionName( oObjPath )
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

'MsgBox "シンボリックリンクを作成しました", vbYes, sPROG_NAME

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

