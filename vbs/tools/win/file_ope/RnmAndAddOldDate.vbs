'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 事前処理
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )      'ConvDate2String()

'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "前日末日時追加＆リネーム"

Dim bIsContinue
bIsContinue = True

Dim lAnswer
lAnswer = MsgBox ( _
                "ファイル/フォルダ名の末尾に前日末日時を付与します。よろしいですか？", _
                vbYesNo, _
                sPROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "実行をキャンセルしました。", vbOKOnly, sPROG_NAME
    bIsContinue = False
End If

Dim cFilePaths

'*******************************************************
'* ファイル/フォルダ名取得
'*******************************************************
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
        Dim objWshShell
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        objWshShell.Run "cmd /c echo.> ""C:\Users\draem_000\Desktop\test.txt""", 0, True
        objWshShell.Run "cmd /c mkdir ""C:\Users\draem_000\Desktop\test2""", 0, True
        cFilePaths.Add "C:\Users\draem_000\Desktop\test.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test2"
    End If
    
    '*** ファイルパスチェック ***
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

'*******************************************************
'* 追加文字列取得
'*******************************************************
If bIsContinue = True Then
    Dim sDateRaw
    Dim sDateStr
    Dim sAddStr
    sDateRaw = Now()
    sDateRaw = DateAdd("d", -1, sDateRaw)
    sDateRaw = Year( sDateRaw ) & "/" & _
               Month( sDateRaw ) & "/" & _
               Day( sDateRaw ) & " " & _
               "17:45:00"
    sDateStr = ConvDate2String( sDateRaw, 1 )
    sAddStr = "_" & sDateStr
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim oFilePath
    For Each oFilePath In cFilePaths
        '*******************************************************
        '* ファイル/フォルダ名判別
        '*******************************************************
        Dim lFileOrFolder '1:ファイル、2:フォルダ、0:エラー（存在しないパス）
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( oFilePath )
        bFileExists = objFSO.FileExists( oFilePath )
        If bFolderExists = False And bFileExists = True Then
            lFileOrFolder = 1 'ファイル
        ElseIf bFolderExists = True And bFileExists = False Then
            lFileOrFolder = 2 'フォルダー
        Else
            lFileOrFolder = 0 'エラー（存在しないパス）
        End If
        
        '*******************************************************
        '* ファイル/フォルダ名変更
        '*******************************************************
        Dim sTrgtDirPath
        Dim sTrgtFileName
        sTrgtDirPath = Mid( oFilePath, 1, InStrRev( oFilePath, "\" ) - 1 )
        sTrgtFileName = Mid( oFilePath, InStrRev( oFilePath, "\" ) + 1, Len( oFilePath ) )
        
        If lFileOrFolder = 1 Then
            If InStr( sTrgtFileName, "." ) > 0 Then
                Dim sTrgtFileBaseName
                Dim sTrgtFileExt
                sTrgtFileExt = Mid( sTrgtFileName, InStrRev( sTrgtFileName, "." ) + 1, Len( sTrgtFileName ) )
                sTrgtFileBaseName = Mid( _
                        sTrgtFileName, _
                        InStrRev( sTrgtFileName, "\" ) + 1, _
                        InStrRev( sTrgtFileName, "." ) - InStrRev( sTrgtFileName, "\" ) - 1 _
                    )
                objFSO.MoveFile _
                    oFilePath, _
                    sTrgtDirPath & "\" & sTrgtFileBaseName & sAddStr & "." & sTrgtFileExt
            Else
                objFSO.MoveFile _
                    oFilePath, _
                    sTrgtDirPath & "\" & sTrgtFileName & sAddStr
            End If
        ElseIf lFileOrFolder = 2 Then
            objFSO.MoveFolder _
                oFilePath, _
                sTrgtDirPath & "\" & sTrgtFileName & sAddStr
        Else
            MsgBox "ファイル/フォルダが不正です。", vbOKOnly, sPROG_NAME
            bIsContinue = False
        End If
        
        If bIsContinue = True Then
            'Do Nothing
        Else
            Exit For
        End If
    Next
    
    Set objFSO = Nothing
Else
    'Do Nothing
End If

'####################################################################
'### インクルード関数
'####################################################################
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

