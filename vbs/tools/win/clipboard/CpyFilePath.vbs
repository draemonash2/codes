'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行
'
' 使い方
'   CpyFilePath.vbs [-d <match_dir_name>] [-l <rem_dir_level>] <file_path1> [<file_path2>]...
'   
'     ex1)
'         CpyFilePath.vbs c\test1.txt c\test2.txt
'             → c\test1.txt<改行>c\test2.txt をコピー
'     ex2)
'         CpyFilePath.vbs -d codes -l 1 c\codes\aaa\bbb\ccc\test.txt c\codes\aaa\bbb\a.txt
'             → bbb\ccc\test.txt<改行>bbb\a.txt をコピー
'     
'     (※) EXECUTION_MODE = 1 にて相対パス置換したい場合は、本スクリプト実行前に
'          sMatchDirNamesRemoveDirLevelに値を設定しておくこと
'          （ex. sMatchDirName = "codes" sRemoveDirLevel = "1"）

'####################################################################
'### 設定
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### インクルード
'####################################################################
Call Include( "C:\codes\vbs\_lib\String.vbs" )

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ファイルパスをコピー"

Dim bIsContinue
bIsContinue = True

Dim cFilePaths
Dim sMatchDirName
Dim sRemoveDirLevel

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim bIsGetDirNameTiming
        Dim bIsGetDirLevelTiming
        bIsGetDirNameTiming = False
        bIsGetDirLevelTiming = False
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            Select Case sArg
                Case "-d"
                    bIsGetDirNameTiming = True
                    bIsGetDirLevelTiming = False
                Case "-l"
                    bIsGetDirLevelTiming = True
                    bIsGetDirNameTiming = False
                Case Else
                    If bIsGetDirNameTiming = True Then
                        sMatchDirName = sArg
                        bIsGetDirNameTiming = False
                    ElseIf bIsGetDirLevelTiming = True Then
                        sRemoveDirLevel = sArg
                        bIsGetDirLevelTiming = False
                    Else
                        cFilePaths.add sArg
                    End If
            End Select
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
        MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
        MsgBox "処理を中断します", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** 相対パスに置換 ***
Dim oFilePath
If sMatchDirName <> "" And sRemoveDirLevel <> "" Then
    Dim cRltvFilePaths
    Set cRltvFilePaths = CreateObject("System.Collections.ArrayList")
    For Each oFilePath In cFilePaths
        Dim sRltvFilePath
        Call ExtractRelativePath(oFilePath, sMatchDirName, CLng(sRemoveDirLevel), sRltvFilePath)
        'Msgbox oFilePath & "：" & sRltvFilePath '★debug
        cRltvFilePaths.Add sRltvFilePath
    Next
    Set cFilePaths = cRltvFilePaths
End If

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

' 外部プログラム インクルード関数
Private Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sOpenFile = objFSO.GetAbsolutePathName( sOpenFile )
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
