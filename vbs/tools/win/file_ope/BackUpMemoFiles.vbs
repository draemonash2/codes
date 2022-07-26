Option Explicit

'<<概要>>
'  指定したフォルダ配下のファイルをバックアップする。
'  
'<<使用方法>>
'  BackUpMemoFiles.vbs <rootdirpath> <backupnum> <backuplogpath>
'  
'<<仕様>>
'  ・<rootdirpath> 配下の中で sEXTRACT_FILE_NAME_PATTERN にマッチするファイルをバックアップする。
'    （_bakフォルダ配下のものは対象外）
'  ・バックアップの仕様は <scriptpath> に準ずる。
'  
'<<依存スクリプト>>
'  ・BackUpMemoFiles.vbs

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= 設定値
'===============================================================================
Const sEXTRACT_FILE_NAME_PATTERN = "\\#memo.*\.xlsm$"
Const sBACKUP_SCRIPT_NAME = "BackUpFile.vbs"

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "ファイル一括バックアップ"

Dim sBakSrcRootDirPath
Dim lBakFileNum
Dim sBakSrcLogPath
If WScript.Arguments.Count >= 3 Then
    sBakSrcRootDirPath = WScript.Arguments(0)
    lBakFileNum = CLng(WScript.Arguments(1))
    sBakSrcLogPath = WScript.Arguments(2)
Else
    WScript.Echo "引数を指定してください。プログラムを中断します。"
    WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sBakScriptPath
sBakScriptPath = objFSO.GetParentFolderName( WScript.ScriptFullName ) & "\" & sBACKUP_SCRIPT_NAME

Dim cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sBakSrcRootDirPath, cFilePaths, 1, "")

Dim oRegExp1
Set oRegExp1 = CreateObject("VBScript.RegExp")
Dim oRegExp2
Set oRegExp2 = CreateObject("VBScript.RegExp")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

oRegExp1.Pattern = sEXTRACT_FILE_NAME_PATTERN
oRegExp1.IgnoreCase = True
oRegExp1.Global = True
oRegExp2.Pattern = "\\_bak" & sEXTRACT_FILE_NAME_PATTERN
oRegExp2.IgnoreCase = True
oRegExp2.Global = True

Dim oMatchResult
Dim vFilePath
For Each vFilePath In cFilePaths
    Set oMatchResult = oRegExp1.Execute(vFilePath)
    If oMatchResult.Count > 0 Then
        Set oMatchResult = oRegExp2.Execute(vFilePath)
        If oMatchResult.Count = 0 Then
            Dim sCmdStr
            sCmdStr = """" & sBakScriptPath & """ """ & vFilePath & """ " & lBakFileNum & " """ & sBakSrcLogPath & """"
            'WScript.Echo sCmdStr
            objWshShell.Run sCmdStr, 0, True
        End If
    End If
Next

'WScript.Echo "バックアップ完了！", vbOKOnly, sSCRIPT_NAME

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
