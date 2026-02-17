'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################

'####################################################################
'### インクルード
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'WinPathToWslPath()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "WSL2のファイルパスをコピー"

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cFilePaths.Count = 0 Then
        MsgBox "ファイルが選択されていません", vbYes, sPROG_NAME
        MsgBox "処理を中断します", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** WindowsパスをWsl2パスに置換 ***
Dim oFilePath
Dim cWsl2FilePaths
Set cWsl2FilePaths = CreateObject("System.Collections.ArrayList")
For Each oFilePath In cFilePaths
    Dim sWsl2FilePath
    sWsl2FilePath = WinPathToWslPath(oFilePath)
    'Msgbox oFilePath & "：" & sWsl2FilePath '★debug
    cWsl2FilePaths.Add sWsl2FilePath
Next
Set cFilePaths = cWsl2FilePaths

'*** クリップボードへコピー ***
If bIsContinue = True Then
    Dim sOutString
    Dim bFirstStore
    bFirstStore = True
    For Each oFilePath In cFilePaths
        If bFirstStore = True Then
            sOutString = oFilePath
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & oFilePath
        End If
    Next
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If

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

