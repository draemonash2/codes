Const sADDIN_DIR_PATH = "C:\codes\vba\excel\AddIns"

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileListCmdClct()
Call Include( "C:\codes\vbs\_lib\Debug.vbs" )       'DebugPrintClct()

'===============================================================================
'= 本処理
'===============================================================================
'フォルダパス一覧取得
Dim cDirPathList
Set cDirPathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sADDIN_DIR_PATH, cDirPathList, 2, "")
'Call DebugPrintClct(cDirPathList) '★デバッグ用★

'コピー元フォルダ判定
Dim sSrcDirPath
sSrcDirPath = ""
Dim vDirPathTmp
For Each vDirPathTmp In cDirPathList
    Dim oRegExp
    Dim sTargetStr
    Dim sSearchPattern
    Set oRegExp = CreateObject("VBScript.RegExp")
    sTargetStr = vDirPathTmp
    sSearchPattern = "\Tmp\d{8}$"
    oRegExp.Pattern = sSearchPattern
    oRegExp.IgnoreCase = True
    oRegExp.Global = True
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count > 0 Then
        sSrcDirPath = vDirPathTmp
        Exit For
    End If
Next
'MsgBox sSrcDirPath: WScript.Quit '★デバッグ用★
If sSrcDirPath = "" Then
    Msgbox "コピー元フォルダが見つからないため、処理を中断します", vbOkOnly, WScript.ScriptName
    WScript.Quit
End If

'コピー元フォルダ内のファイルリスト取得
Dim cSrcFilePathList
Set cSrcFilePathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sSrcDirPath, cSrcFilePathList, 1, "")
'Call DebugPrintClct(cSrcFilePathList) '★デバッグ用★

'コピー元フォルダ内のファイルをコピー先フォルダへコピー
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Dim sDebugMsg '★デバッグ用★
'sDebugMsg = "" '★デバッグ用★
Dim vSrcFilePath
Dim sDstDirPath
sDstDirPath = sADDIN_DIR_PATH & "\MyExcelAddin.bas\"
For Each vSrcFilePath In cSrcFilePathList
    objFSO.CopyFile vSrcFilePath, sDstDirPath
    'sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath & "→" & sDstDirPath '★デバッグ用★
Next
'MsgBox sDebugMsg: WScript.Quit '★デバッグ用★

'コピー元フォルダ削除
objFSO.DeleteFolder sSrcDirPath, True

Msgbox "更新完了！", vbOkOnly, WScript.ScriptName

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function


