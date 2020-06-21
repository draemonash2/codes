'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "Svn リビジョンをコピー"

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
        cFilePaths.Add "C:\codes_sample\c\FreeRTOSV7.1.1\Source\croutine.c"
        cFilePaths.Add "C:\codes_sample\c\FreeRTOSV7.1.1\Source\include\test.txt"
        cFilePaths.Add "C:\codes_sample\c\FreeRTOSV7.1.1\Source\test.txt"
        cFilePaths.Add "C:\codes_sample\c\FreeRTOSV7.1.1\Source\nothing"
        cFilePaths.Add "C:\codes_sample\c\FreeRTOSV7.1.1\Source"
        cFilePaths.Add "C:\codes_sample\_mergetest"
        cFilePaths.Add "C:\codes_sample\_mergetest\01\Add_practice09.c"
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

'*** クリップボードへコピー ***
If bIsContinue = True Then
    Dim sOutString
    sOutString = ""
    For Each oFilePath In cFilePaths
        Dim sRevision
        sRevision = GetSvnRevision( oFilePath )
        If sOutString = "" Then
            sOutString = sRevision
        Else
            sOutString = sOutString & vbNewLine & sRevision
        End If
    Next
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If

'リビジョンを取得する
Private Function GetSvnRevision( _
    ByVal sTrgtPath _
)
    Const sEXT_KEYWORD_REV = "Last Changed Rev: "
    Const sCMD_RDRCT_FILE_NAME = "svn_info.log"
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim sTmpDirPath
    sTmpDirPath = objFSO.GetSpecialFolder(2) 'テンポラリフォルダ
    Dim sCmdRdrctFilePath
    sCmdRdrctFilePath = sTmpDirPath & "\" & sCMD_RDRCT_FILE_NAME
    
    'svn info 実行
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c svn info """ & sTrgtPath & """ > " & sCmdRdrctFilePath & " 2>&1", 0, True
    
    'svn info 実行結果取得
    Dim sRevision
    Dim objTxtFile
    Set objTxtFile = objFSO.OpenTextFile( sCmdRdrctFilePath, 1, True)
    Do Until objTxtFile.AtEndOfStream
        Dim sLine
        sLine = objTxtFile.ReadLine
        If InStr( sLine, sEXT_KEYWORD_REV ) > 0 Then
            sRevision = Mid( sLine, Len( sEXT_KEYWORD_REV ) + 1, Len( sLine ) )
            Exit Do
        End If
    Loop
    objTxtFile.Close
    
    GetSvnRevision = sRevision
End Function
