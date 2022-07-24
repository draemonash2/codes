Option Explicit

'<<概要>>
'  指定したフォルダ配下のファイルをバックアップする。
'  
'<<使用方法>>
'  BackUpFilesAll.vbs <scriptpath> <rootdirpath> <filepathpattern> <backupnum> <backuplogpath>
'  
'<<仕様>>
'  ・<rootdirpath> 配下の中で <filepathpattern> にマッチするファイルをバックアップする。
'    （_bakフォルダ配下のものは対象外）
'  ・バックアップの仕様は <scriptpath> に準ずる。

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= 設定値
'===============================================================================

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "ファイル一括バックアップ"

Dim sBakScriptPath
Dim sBakSrcRootDirPath
Dim sExtractFileNamePattern
Dim lBakFileNum
Dim sBakSrcLogPath
If WScript.Arguments.Count >= 5 Then
    sBakScriptPath = WScript.Arguments(0)
    sBakSrcRootDirPath = WScript.Arguments(1)
    sExtractFileNamePattern = WScript.Arguments(2)
    lBakFileNum = CLng(WScript.Arguments(3))
    sBakSrcLogPath = WScript.Arguments(4)
Else
    WScript.Echo "引数を指定してください。プログラムを中断します。"
    WScript.Quit
End If

Dim cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sBakSrcRootDirPath, cFilePaths, 1, "")

Dim oRegExp1
Set oRegExp1 = CreateObject("VBScript.RegExp")
Dim oRegExp2
Set oRegExp2 = CreateObject("VBScript.RegExp")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

oRegExp1.Pattern = sExtractFileNamePattern & "$"
oRegExp1.IgnoreCase = True
oRegExp1.Global = True
oRegExp2.Pattern = "\\_bak\\" & sExtractFileNamePattern & "$"
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
