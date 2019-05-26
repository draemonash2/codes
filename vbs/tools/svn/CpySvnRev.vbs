'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "Svn リビジョンをコピー"

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

'*** クリップボードへコピー ***
If bIsContinue = True Then
    Dim sOutString
    Dim objTxtFile
    Dim sPrgNo
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
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
    Set objFSO = Nothing
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If

'リビジョンを取得する
Private Function GetSvnRevision( _
    ByVal sTrgtPath _
)
    Const sEXT_KEYWORD_REV = "Last Changed Rev: "
    Const sEXT_KEYWORD_NOTSVN = "is not a working copy"
    
    Dim sExeCmd
    sExeCmd = "svn info " & sTrgtPath & "|find """ & sEXT_KEYWORD_REV & """"
    MsgBox sExeCmd
    
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objExecResult
    Set objExecResult = objWshShell.Exec( sExeCmd )
    
    '終了待ち
    Do While objExecResult.Status = 0
        WScript.Sleep 100
    Loop
    
    Dim sRet
    sRet = ""
    MsgBox objExecResult.ExitCode
    If objExecResult.ExitCode = 0 Then
        Dim bIsError
        bIsError = True
        
        Dim sLine
        Do Until objExecResult.StdOut.AtEndOfStream
            sLine = objExecResult.StdOut.ReadLine
        Loop
        MsgBox sLine
        If InStr( sLine, sEXT_KEYWORD_REV ) > 0 Then
            bIsError = False
        ElseIf InStr( sLine, sEXT_KEYWORD_NOTSVN ) > 0 Then
            bIsError = False
        Else
            'Do Nothing
        End If
        
        Dim sRevision
        If bIsError = False Then
            sRevision = Mid( sLine, Len( sEXT_KEYWORD_REV ) + 1, Len( sLine ) )
            sRet = sRevision
        Else
            'MsgBox "svn コマンドが正常に実行できませんでした。"
        End If
    Else
        'MsgBox "異常終了！"
    End If
    GetSvnRevision = sRet
End Function
