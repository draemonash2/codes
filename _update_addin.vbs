Const sAddinDirPath = "C:\codes\vba\excel\AddIns"

'フォルダパス一覧取得
Dim cDirPathList
Set cDirPathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sAddinDirPath, cDirPathList, 2, "")

''★デバッグ用★
'Dim sDebugMsg
'sDebugMsg = ""
'Dim vDirPath
'For Each vDirPath In cDirPathList
'    sDebugMsg = sDebugMsg & vbNewLine & vDirPath
'Next
'MsgBox sDebugMsg
'WScript.Quit

'コピー元フォルダ判定
Dim sSrcDirPath
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

''★デバッグ用★
'MsgBox sSrcDirPath
'WScript.Quit

'コピー元フォルダ内のファイルリスト取得
Dim cSrcFilePathList
Set cSrcFilePathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sSrcDirPath, cSrcFilePathList, 1, "")

''★デバッグ用★
'Dim sDebugMsg
'sDebugMsg = ""
'For Each vSrcFilePath In cSrcFilePathList
'    sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath
'Next
'MsgBox sDebugMsg
'WScript.Quit

'コピー元フォルダ内のファイルをコピー先フォルダへコピー
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Dim sDebugMsg '★デバッグ用★
'sDebugMsg = "" '★デバッグ用★
Dim vSrcFilePath
Dim sDstDirPath
sDstDirPath = sAddinDirPath & "\MyExcelAddin.bas\"
For Each vSrcFilePath In cSrcFilePathList
    objFSO.CopyFile vSrcFilePath, sDstDirPath
    'sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath & "→" & sDstDirPath '★デバッグ用★
Next
'MsgBox sDebugMsg '★デバッグ用★
'WScript.Quit '★デバッグ用★

'コピー元フォルダ削除
objFSO.DeleteFolder sSrcDirPath, True

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する(Collection,Dirコマンド版)
' = 引数    sTrgtDir        String      [in]    対象フォルダ
' = 引数    cFileList       Collections [out]   ファイル/フォルダパス一覧
' = 引数    lFileListType   Long        [in]    取得する一覧の形式
' =                                                 0：両方
' =                                                 1:ファイル
' =                                                 2:フォルダ
' =                                                 それ以外：格納しない
' = 引数    sFileExtStr     String      [in]    取得するファイルの拡張子
' =                                                 ex1) ""
' =                                                 ex2) "*"
' =                                                 ex3) "*.c"
' =                                                 ex4) "*.txt *.log *.csv"
' = 戻値    なし
' = 覚書    ・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' = 覚書    ・sFileExtStrはファイル指定時のみ有効
' = 依存    なし
' = 所属    FileSystem.vbs
' ==================================================================
Public Function GetFileListCmdClct( _
    ByVal sTrgtDir, _
    ByRef cFileList, _
    ByVal lFileListType, _
    ByVal sFileExtStr _
)
    Dim objFSO  'FileSystemObjectの格納先
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    
    'Dir コマンド実行（出力結果を一時ファイルに格納）
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = WScript.CreateObject( "WScript.Shell" ).CurrentDirectory & "\Dir.tmp"
    Dim sTrgtDirStr
    If sFileExtStr = "" Then
        sTrgtDirStr = """" & sTrgtDir & """"
    Else
        Dim vFileExtentions
        vFileExtentions = Split(sFileExtStr, " ")
        Dim lSplitIdx
        For lSplitIdx = 0 To UBound(vFileExtentions)
            If sTrgtDirStr = "" Then
                sTrgtDirStr = """" & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            Else
                sTrgtDirStr = sTrgtDirStr & " """ & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            End If
        Next
    End If
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile
    Dim sTextAll
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
            Dim vFileList
            vFileList = Split( sTextAll, vbNewLine )
            Dim sFilePath
            For Each sFilePath In vFileList
                cFileList.add sFilePath
            Next
            objFile.Close
        Else
            WScript.Echo "ファイルが開けません: " & Err.Description
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        WScript.Echo "エラー " & Err.Description
    End If  
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error Goto 0
End Function
'   Call Test_GetFileListCmdClct()
    Private Sub Test_GetFileListCmdClct()
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim sCurDir
        sCurDir = "C:\codes"
        
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        Call GetFileListCmdClct( sCurDir, cFileList, 1, "*.c *.h" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 1, "*.h" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 1, "" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 2, "" )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox sOutput
    End Sub
