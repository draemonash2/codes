'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path>

'★TODO★：ZIP ファイル以外の圧縮動作確認

'####################################################################
'### 事前処理
'####################################################################
Dim cAcceptFileFormats
Set cAcceptFileFormats = CreateObject("System.Collections.ArrayList")

Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" ) 'GetNotExistPath()

'####################################################################
'### 設定
'####################################################################
'7-Zip 16.04 圧縮可能形式（/7-ZipPortable/App/7-Zip/7-zip.chm より引用）
'                      [FileExt]    [Format]
cAcceptFileFormats.Add "7z"       ' 7z
cAcceptFileFormats.Add "bz2"      ' BZIP2
cAcceptFileFormats.Add "bzip2"    ' BZIP2
cAcceptFileFormats.Add "tbz2"     ' BZIP2
cAcceptFileFormats.Add "tbz"      ' BZIP2
cAcceptFileFormats.Add "gz"       ' GZIP
cAcceptFileFormats.Add "gzip"     ' GZIP
cAcceptFileFormats.Add "tgz"      ' GZIP
cAcceptFileFormats.Add "tar"      ' TAR
cAcceptFileFormats.Add "wim"      ' WIM
cAcceptFileFormats.Add "swm"      ' WIM
cAcceptFileFormats.Add "xz"       ' XZ
cAcceptFileFormats.Add "txz"      ' XZ
cAcceptFileFormats.Add "zip"      ' ZIP
cAcceptFileFormats.Add "zipx"     ' ZIP
cAcceptFileFormats.Add "jar"      ' ZIP
cAcceptFileFormats.Add "xpi"      ' ZIP
cAcceptFileFormats.Add "odt"      ' ZIP
cAcceptFileFormats.Add "ods"      ' ZIP
cAcceptFileFormats.Add "docx"     ' ZIP
cAcceptFileFormats.Add "xlsx"     ' ZIP
cAcceptFileFormats.Add "epub"     ' ZIP
Const INITIAL_FILE_EXT = "zip"

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "7-Zip で圧縮"

Dim bIsContinue
bIsContinue = True

Dim cSelectedPaths

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\aa"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\b b"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\d.txt"
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'********************
'*** 圧縮形式選択 ***
'********************
If bIsContinue = True Then
    Dim bIsReEnter
    bIsReEnter = False
    Dim sAcceptFileFormatsStr
    Dim sAcceptFileFormat
    sAcceptFileFormatsStr = ""
    For Each sAcceptFileFormat In cAcceptFileFormats
        sAcceptFileFormatsStr = sAcceptFileFormatsStr & vbNewLine & sAcceptFileFormat
    Next
    Do
        Dim sArchiveFileExt
        sArchiveFileExt = InputBox( _
                            "以下の中から圧縮形式を選択して入力してください。" & vbNewLine & _
                            sAcceptFileFormatsStr & vbNewLine, _
                            sPROG_NAME, _
                            INITIAL_FILE_EXT _
                        )
        If sArchiveFileExt = "" Then
            MsgBox "実行をキャンセルしました。", vbOKOnly, sPROG_NAME
            MsgBox "処理を中断します。", vbYes, sPROG_NAME
            bIsReEnter = False
            bIsContinue = False
        Else
            Dim bIsExist
            bIsExist = False
            For Each sAcceptFileFormat In cAcceptFileFormats
                If sAcceptFileFormat = sArchiveFileExt Then
                    bIsExist = True
                Else
                    'Do Nothing
                End If
            Next
            If bIsExist = True Then
                bIsReEnter = False
            Else
                MsgBox "対応する圧縮形式ではありません。" & vbNewLine & vbNewLine & sArchiveFileExt, vbOKOnly, sPROG_NAME
                bIsReEnter = True
            End If
            bIsContinue = True
        End If
    Loop While bIsReEnter = True
Else
    'Do Nothing
End If

'************************
'*** 対象ファイル選定 ***
'************************
'*** ファイル選定 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If bIsContinue = True Then
    Dim cTrgtPaths
    Set cTrgtPaths = CreateObject("System.Collections.ArrayList")
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        If bFolderExists = False And bFileExists = True Then
            cTrgtPaths.Add sSelectedPath
        ElseIf bFolderExists = True And bFileExists = False Then
            cTrgtPaths.Add sSelectedPath
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, sPROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cTrgtPaths.Count = 0 Then
        MsgBox "対象となるファイル/フォルダが存在しません。", vbYes, sPROG_NAME
        MsgBox "処理を中断します。", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'****************
'*** 実行確認 ***
'****************
If bIsContinue = True Then
    Dim sTrgtPath
    Dim sTrgtPathsStr
    sTrgtPathsStr = ""
    For Each sTrgtPath In cTrgtPaths
        If sTrgtPathsStr = "" Then
            sTrgtPathsStr = sTrgtPath
        Else
            sTrgtPathsStr = sTrgtPathsStr & vbNewLine & sTrgtPath
        End If
    Next
    Dim lAnswer
    lAnswer = MsgBox ( _
                    "以下を【圧縮】して、選択ファイルと同じフォルダに格納します。" & vbNewLine & _
                    "よろしいですか？" & vbNewLine & _
                    vbNewLine & _
                    "<<圧縮形式>>" & vbNewLine & _
                    sArchiveFileExt & vbNewLine & _
                    vbNewLine & _
                    "<<対象ファイル/フォルダパス(※)>>" & vbNewLine & _
                    sTrgtPathsStr & vbNewLine & _
                    vbNewLine & _
                    "(※) それぞれのファイル/フォルダが圧縮されます！" & vbNewLine & _
                    "     一つの圧縮ファイルになる訳ではありません！", _
                    vbYesNo, _
                    sPROG_NAME _
                )
    If lAnswer = vbYes Then
        'Do Nothing
    Else
        MsgBox "実行をキャンセルしました。", vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'****************
'*** 圧縮実行 ***
'****************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
If bIsContinue = True Then
    Dim sExePath
    sExePath = objWshShell.Environment("System").Item("MYPATH_7Z")
    If sExePath = "" then
        MsgBox "環境変数が設定されていません。" & vbNewLine & "処理を中断します。", vbYes, sPROG_NAME
        WScript.Quit
    End If
    
    For Each sTrgtPath In cTrgtPaths
        Dim sArchiveFilePath
        Dim bRet
        Dim lAddedPathType
        bRet = GetNotExistPath( sTrgtPath & "." & sArchiveFileExt, sArchiveFilePath, lAddedPathType )
        Dim sExecCmd
        sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtPath & """"
        objWshShell.Run sExecCmd, 1, True
    Next
    MsgBox "圧縮が完了しました。", vbOKOnly, sPROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
