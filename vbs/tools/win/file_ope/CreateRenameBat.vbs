'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const OUTPUT_BAT_FILE_BASE_NAME = "rename"

'####################################################################
'### 事前処理
'####################################################################
Call Include( "C:\codes\vbs\_lib\String.vbs" ) 'LenByte()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "リネーム用バッチファイル出力"

Dim bIsContinue
bIsContinue = True

Dim objFSO
Dim sOutputBatDirPath
Dim cFilePaths

If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Dim sDefaultPath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sOutputBatDirPath = InputBox( "ファイルパスを指定してください", sPROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sOutputBatDirPath = WScript.Env("Current")
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sOutputBatDirPath = "C:\Users\draem_000\Desktop\test"
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

'*** ファイルパスからファイル名取り出し ***
If bIsContinue = True Then
    Dim cFileNames
    Set cFileNames = CreateObject("System.Collections.ArrayList")
    Dim oFilePath
    For Each oFilePath In cFilePaths
        cFileNames.Add Mid( oFilePath, InStrRev( oFilePath, "\" ) + 1, Len( oFilePath ) )
    Next
Else
    'Do Nothing
End If

'*** 最大の文字列長を取得 ***
If bIsContinue = True Then
    Dim lFileNameLenMax
    lFileNameLenMax = 0
    Dim oFileName
    For Each oFileName In cFileNames
        Dim lTrgtFileNameLen
        lTrgtFileNameLen = LenByte( oFileName )
        If lTrgtFileNameLen > lFileNameLenMax Then
            lFileNameLenMax = lTrgtFileNameLen
        Else
            'Do Nothing
        End If
    Next
Else
    'Do Nothing
End If

'*** リネーム前ファイルリスト出力 ***
If bIsContinue = True Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objTxtFile
    Dim sBakFilePath
    sBakFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & "_bak.txt"
    Set objTxtFile = objFSO.OpenTextFile( sBakFilePath, 2, True)
    For Each oFileName In cFileNames
        objTxtFile.WriteLine oFileName
    Next
    objTxtFile.Close
    Set objTxtFile = Nothing
Else
    'Do Nothing
End If

'*** リネーム用バッチファイル出力 ***
If bIsContinue = True Then
    Dim sBatFilePath
    sBatFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & ".bat"
    Set objTxtFile = objFSO.OpenTextFile( sBatFilePath, 2, True)
    For Each oFileName In cFileNames
        objTxtFile.WriteLine _
            "rename " & _
            """" & oFileName & """" & _
            String(lFileNameLenMax - LenByte( oFileName ) + 1, " ") & _
            """" & oFileName & """"
    Next
    objTxtFile.WriteLine "pause"
    objTxtFile.Close
    Set objTxtFile = Nothing
    
    MsgBox OUTPUT_BAT_FILE_BASE_NAME & ".bat を出力しました。"
Else
    'Do Nothing
End If

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

