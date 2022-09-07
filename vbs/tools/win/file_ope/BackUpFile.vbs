Option Explicit

'<<概要>>
'  指定したファイルをバックアップする。
'  
'<<使用方法>>
'  BackUpFile.vbs <filepath> [<backupnum>] [<logfilepath>]
'  
'<<仕様>>
'  ・ファイルを指定すると現在時刻を付与したバックアップファイルを作成する。
'  ・同じファイル名のものが存在していたら、アルファベットを付与したバックアップファイルを作成する。
'     ex. 211201a, 211202b, …
'  ・第二引数に指定されたバックアップ数分ファイルがたまったら、古いものから削除する。
'  ・実行結果は第三引数に指定されたログファイルに出力する。
'  ・前回バックアップ時から更新されていない場合、バックアップしない。
'  
'<<注意事項>>
'  ・バックアップ対象はファイルのみ。
'  ・以下を全て満たす場合、新しいファイルが更新されていくため要注意。
'      - バックアップファイルの接尾辞が"z"となっているファイルがある (ex. file.txt.#b#211122z.txt)
'  ・以下の理由で最新/最古バックアップファイル判定に更新日時を用いない。あくまで
'    バックアップした日を示すファイル名で判断する。
'      誤って古いバックアップファイルを更新してしまった場合、ファイル名上は
'      日付が古いのに更新日時が新しいファイルができてしまう。
'      更新日時をもとに判定すると、上記のファイルが削除されず、残ってしまうため。

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()
                                                            'CreateDirectry()
                                                            'GetFileInfo()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "ファイルバックアップ"
Const sBAK_DIR_NAME = "_#b#"
Const sBAK_FILE_SUFFIX = "#b#"
Const lBAK_FILE_NUM_DEFAULT = 30

'===============================================================================
'= 本処理
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= メイン関数
'===============================================================================
Public Sub Main()
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim sBakSrcFilePath
    Dim lBakFileNumMax
    Dim sBakLogFilePath
    If cArgs.Count >= 3 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = CLng(cArgs(1))
        sBakLogFilePath = cArgs(2)
    ElseIf cArgs.Count >= 2 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = CLng(cArgs(1))
        sBakLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & objFSO.GetBaseName(WScript.ScriptName) & ".log"
    ElseIf cArgs.Count >= 1 Then
        sBakSrcFilePath = cArgs(0)
        lBakFileNumMax = lBAK_FILE_NUM_DEFAULT
        sBakLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & objFSO.GetBaseName(WScript.ScriptName) & ".log"
    Else
        WScript.Echo "引数を指定してください。プログラムを中断します。"
        Exit Sub
    End If
    'MsgBox sBakSrcFilePath & vbNewLine & lBakFileNumMax & vbNewLine & sBakLogFilePath
    
    Dim objLogFile
    Set objLogFile = objFSO.OpenTextFile(sBakLogFilePath, 8, True) 'AddWrite
    
    '****************
    '*** 事前準備 ***
    '****************
    '対象ファイル情報取得
    Dim sBakSrcParDirPath
    Dim sBakSrcFileName
    Dim sBakSrcFileBaseName
    Dim sBakSrcFileExt
    Dim sDateSuffix
    sBakSrcParDirPath = objFSO.GetParentFolderName( sBakSrcFilePath )
    sBakSrcFileName = objFSO.GetFileName( sBakSrcFilePath )
    sBakSrcFileBaseName = objFSO.GetBaseName( sBakSrcFilePath )
    sBakSrcFileExt = objFSO.GetExtensionName( sBakSrcFilePath )
    sDateSuffix = ConvDate2String(Now(),2)
    'MsgBox sBakSrcParDirPath & vbNewLine & sBakSrcFileName & vbNewLine & sBakSrcFileBaseName & vbNewLine & sBakSrcFileExt & vbNewLine & sDateSuffix
    
    '拡張子有無チェック
    Dim bExistsExt
    If ( (sBakSrcFileBaseName <> "") And (sBakSrcFileExt <> "") ) Then
        bExistsExt = True
    Else
        bExistsExt = False
    End If
    
    'バックアップファイル情報作成
    Dim sBakDstDirPath
    Dim sBakDstPathBase
    sBakDstDirPath = sBakSrcParDirPath & "\" & sBAK_DIR_NAME
    sBakDstPathBase = sBakDstDirPath & "\" & sBakSrcFileName & "." & sBAK_FILE_SUFFIX
    
    '****************************
    '*** ファイルバックアップ ***
    '****************************
    'バックアップフォルダ作成
    Call CreateDirectry( sBakDstDirPath )
    
    'ファイル一覧取得
    Dim cFileList
    Set cFileList = CreateObject("System.Collections.ArrayList")
    Call GetFileListCmdClct( sBakDstDirPath, cFileList, 1, "*")
    
    '既存の最新ファイル探索
    Dim sBakDstFilePathLatest  '既存の最新バックアップファイル
    sBakDstFilePathLatest = ""
    Dim sFilePath
    For Each sFilePath In cFileList
        If ( InStr(sFilePath, sBakDstPathBase) > 0 ) Then
            sBakDstFilePathLatest = sFilePath
        End If
    Next
    Set cFileList = Nothing
    
    'バックアップファイル名確定
    Dim sBakDstFilePath
    '既存のバックアップファイルが存在し、同じ日付のバックアップファイルが存在する場合
    If sBakDstFilePathLatest <> "" And _
       InStr(sBakDstFilePathLatest, sBakDstPathBase & sDateSuffix) > 0 Then
        Dim sTailChar
        If bExistsExt = True Then
            sTailChar = Right( objFSO.GetBaseName( sBakDstFilePathLatest ), 1)
        Else
            sTailChar = Right( sBakDstFilePathLatest, 1)
        End If
        Dim lBakDstAlphaIdx
        If Asc(sTailChar) >= Asc("a") And Asc(sTailChar) < Asc("z") Then
            lBakDstAlphaIdx = Asc(sTailChar) + 1
        ElseIf Asc(sTailChar) = Asc("z") Then
            lBakDstAlphaIdx = Asc(sTailChar)
        ElseIf Asc(sTailChar) >= Asc("0") And Asc(sTailChar) <= Asc("9") Then
            lBakDstAlphaIdx = Asc("a")
        Else
            objLogFile.WriteLine "不正なバックアップファイルが見つかりました。"
            objLogFile.WriteLine "  " & sBakDstFilePathLatest
            objLogFile.WriteLine "プログラムを中断します。"
            Exit Sub
        End If
        If bExistsExt = True Then
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx) & "." & sBakSrcFileExt
        Else
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx)
        End If
    Else
        If bExistsExt = True Then
            sBakDstFilePath = sBakDstPathBase & sDateSuffix & "." & sBakSrcFileExt
        Else
            sBakDstFilePath = sBakDstPathBase & sDateSuffix
        End If
    End If
    'objLogFile.WriteLine sBakDstFilePath & " : " & sBakDstFilePathLatest
    'WScript.Echo sBakDstFilePath & " : " & sBakDstFilePathLatest
    'Exit Sub
    
    '更新日時取得
    Dim vDateLastModifiedLatestBk
    Dim vDateLastModifiedTrgt
    Dim bRet
    bRet = GetFileInfo( sBakDstFilePathLatest, 11, vDateLastModifiedLatestBk)
    bRet = GetFileInfo( sBakSrcFilePath, 11, vDateLastModifiedTrgt)
    
    '既存のバックアップファイル未存在 or 更新されている場合
    Dim sLogMsg
    sLogMsg = ""
    If ( sBakDstFilePathLatest = "" ) Or _
       ( ( sBakDstFilePathLatest <> "" ) And ( vDateLastModifiedTrgt > vDateLastModifiedLatestBk ) ) Then
        'ファイルバックアップ
        objFSO.CopyFile sBakSrcFilePath, sBakDstFilePath, True
        sLogMsg = "[Success] " & sBakSrcFilePath & " -> " & sBakDstFilePath & "."
    Else
        '前回バックアップ時から更新されていない場合、バックアップせず処理を中断する
        objLogFile.WriteLine "[Skip]    " & sBakSrcFilePath & "."
        Exit Sub
    End If
    
    '************************
    '*** 古いファイル削除 ***
    '************************
    'ファイルリスト取得
    Dim cFileListAll
    Set cFileListAll = CreateObject("System.Collections.ArrayList")
    Call GetFileListCmdClct( sBakDstDirPath, cFileListAll, 1, "*")
    Set cFileList = CreateObject("System.Collections.ArrayList")
    For Each sFilePath in cFileListAll
        If InStr(sFilePath, sBakDstPathBase) > 0 Then
            cFileList.Add sFilePath
        End If
    Next
    
    'バックアップファイル削除
    Dim lBakFileNum
    Dim lDelFileNum
    lBakFileNum = cFileList.Count
    lDelFileNum = 0
    For Each sFilePath In cFileList
        If lBakFileNum > lBakFileNumMax Then
            'objFSO.DeleteFile sFilePath, True
            Call MoveToTrushBox(objFSO, sFilePath)
            lDelFileNum = lDelFileNum + 1
        End If
        lBakFileNum = lBakFileNum - 1
    Next
    Set cFileList = Nothing
    
    If lDelFileNum > 0 Then
        objLogFile.WriteLine sLogMsg & " " & lDelFileNum & " files deleted."
    Else
        objLogFile.WriteLine sLogMsg
    End If
    
    'objLogFile.WriteLine "バックアップ完了！", vbOKOnly, sSCRIPT_NAME
    
    objLogFile.Close
End Sub

'===============================================================================
'= 内部関数
'===============================================================================
Private Function MoveToTrushBox(ByRef objFSO, ByVal sTrgtPath)
    If objFSO.FileExists(sTrgtPath) Then
        CreateObject("Shell.Application").Namespace(10).movehere sTrgtPath
        Do While objFSO.FileExists(sTrgtPath) Or objFSO.FolderExists(sTrgtPath)
            '削除処理は非同期で進行するため、削除中にスクリプトが終了すると削除処理は中断される。
            '削除対象が削除されるまで待機する。
            WScript.sleep(100)
        Loop
        MoveToTrushBox = True
    Else
        MoveToTrushBox = False
    End If
End Function

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim sDesktopPath
    sDesktopPath = objWshShell.SpecialFolders("Desktop")
    
    MsgBox "=== test start ==="
    
    Dim sTrgtFilePath
    Dim sTrgtFilePathOrg
    Dim sBakDirPath
    Dim sBakLogName
    Dim objTxtFile
    Select Case lTestCase
        Case 1
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            sBakLogName = sDesktopPath & "\backup_test.log"
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            cArgs.Add 5
            cArgs.Add sBakLogName
            
            Call Main()
            MsgBox "1 バックアップ生成後(無印追加)"
            
            Dim objDummyFile
            Set objDummyFile = objFSO.OpenTextFile(sDesktopPath & "\" & sBAK_DIR_NAME & "\dummy_file.txt", 8, True)
            objDummyFile.WriteLine "a"
            objDummyFile.Close
            
            Call Main()
            MsgBox "2 バックアップ生成後(変化なし)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "3 バックアップ生成後(a追加)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "4 バックアップ生成後(b追加)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "5 バックアップ生成後(c追加)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "6 バックアップ生成後(d追加)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "7 バックアップ生成後(e追加 無印削除)"
            
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "8 バックアップ生成後(f追加 a削除)"
            
            cArgs(1) = 2
            Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 8, True)
            objTxtFile.WriteLine "aa"
            objTxtFile.Close
            Call Main()
            MsgBox "9 バックアップ生成後(g追加 b,c,d,e削除)"
        Case 2
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            cArgs.Add 5
            Call Main()
        Case 3
            sTrgtFilePath = sDesktopPath & "\backup_test.txt"
            sTrgtFilePathOrg = sDesktopPath & "\backup_test_org.txt"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            
            If objFSO.FileExists(sTrgtFilePathOrg) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePathOrg, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
            End If
            objFSO.CopyFile sTrgtFilePathOrg, sTrgtFilePath, True
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath
            Call Main()
        Case 4
            Dim sTrgtFilePath1
            Dim sTrgtFilePath2
            Dim sTrgtFilePath3
            Dim sTrgtFilePath4
            sTrgtFilePath1 = sDesktopPath & "\backup_test.txt"
            sTrgtFilePath2 = sDesktopPath & "\.backup_test.txt"
            sTrgtFilePath3 = sDesktopPath & "\backup_test"
            sTrgtFilePath4 = sDesktopPath & "\.backup_test"
            sBakDirPath = sDesktopPath & "\" & sBAK_DIR_NAME
            sBakLogName = sDesktopPath & "\backup_test.log"
            
            If objFSO.FileExists(sTrgtFilePath1) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath1, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath2) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath2, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath3) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath3, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FileExists(sTrgtFilePath4) Then
                'Do Nothing
            Else
                Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath4, 8, True)
                objTxtFile.WriteLine "a"
                objTxtFile.Close
                Set objTxtFile = Nothing
            End If
            If objFSO.FolderExists( sBakDirPath ) Then
                objFSO.DeleteFolder sBakDirPath, True
            End If
            
            cArgs.Add sTrgtFilePath1
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath2
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath3
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
            
            cArgs.Add sTrgtFilePath4
            cArgs.Add 5
            cArgs.Add sBakLogName
            Call Main()
            cArgs.Clear
        Case Else
            Call Main()
    End Select
    
    MsgBox "=== test finished ==="
End Sub '}}}

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}
