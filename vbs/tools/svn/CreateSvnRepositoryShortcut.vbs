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

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim asSplitWord
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
'   Call Test_ExtractTailWord()
    Private Sub Test_ExtractTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
        Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
        MsgBox Result
    End Sub
