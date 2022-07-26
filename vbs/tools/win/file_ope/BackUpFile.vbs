Option Explicit

'<<概要>>
'  指定したファイルをバックアップする。
'  
'<<使用方法>>
'  BackUpFile.vbs <filepath> <backupnum> <logfilepath>
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
'      - バックアップファイルの接尾辞が"z"となっているファイルがある (ex. file_#b#211122z.txt)
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
Const sBAK_DIR_NAME = "_bak"
Const sBAK_FILE_SUFFIX = "#b#"

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "ファイルバックアップ"
Dim sBakSrcFilePath
Dim lBakFileNumMax
Dim sBakLogFilePath
If WScript.Arguments.Count >= 3 Then
    sBakSrcFilePath = WScript.Arguments(0)
    lBakFileNumMax = CLng(WScript.Arguments(1))
    sBakLogFilePath = WScript.Arguments(2)
Else
    WScript.Echo "引数を指定してください。プログラムを中断します。"
    WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objLogFile
Set objLogFile = objFSO.OpenTextFile(sBakLogFilePath, 8, True) 'AddWrite

'****************
'*** 事前準備 ***
'****************
'対象ファイル情報取得
Dim sBakSrcParDirPath
Dim sBakSrcFileBaseName
Dim sBakSrcFileExt
Dim sDateSuffix
sBakSrcParDirPath = objFSO.GetParentFolderName( sBakSrcFilePath )
sBakSrcFileBaseName = objFSO.GetBaseName( sBakSrcFilePath )
sBakSrcFileExt = objFSO.GetExtensionName( sBakSrcFilePath )
sDateSuffix = ConvDate2String(Now(),2)

'バックアップファイル情報作成
Dim sBakDstDirPath
Dim sBakDstPathBase
sBakDstDirPath = sBakSrcParDirPath & "\" & sBAK_DIR_NAME
sBakDstPathBase = sBakDstDirPath & "\" & sBakSrcFileBaseName & "_" & sBAK_FILE_SUFFIX

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
    If ( ( InStr(sFilePath, sBakDstPathBase) > 0 ) And _
       (objFSO.GetExtensionName(sFilePath) = sBakSrcFileExt) ) Then
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
    sTailChar = Right( objFSO.GetBaseName( sBakDstFilePathLatest ), 1)
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
        WScript.Quit
    End If
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx) & "." & sBakSrcFileExt
Else
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & "." & sBakSrcFileExt
End If
'objLogFile.WriteLine sBakDstFilePath & " : " & sBakDstFilePathLatest
'WScript.Quit

'更新日時取得
Dim vDateLastModifiedLatestBk
Dim vDateLastModifiedTrgt
Dim bRet
bRet = GetFileInfo( sBakDstFilePathLatest, 11, vDateLastModifiedLatestBk)
bRet = GetFileInfo( sBakSrcFilePath, 11, vDateLastModifiedTrgt)

'既存のバックアップファイル未存在 or 更新されている場合
If ( sBakDstFilePathLatest = "" ) Or _
   ( ( sBakDstFilePathLatest <> "" ) And ( vDateLastModifiedTrgt > vDateLastModifiedLatestBk ) ) Then
    'ファイルバックアップ
    objFSO.CopyFile sBakSrcFilePath, sBakDstFilePath, True
    objLogFile.WriteLine "[Success] " & sBakSrcFilePath & " -> " & sBakDstFilePath
Else
    '前回バックアップ時から更新されていない場合、バックアップせず処理を中断する
    objLogFile.WriteLine "[Skip]    " & sBakSrcFilePath
    WScript.Quit
End If

'************************
'*** 古いファイル削除 ***
'************************
'ファイルリスト取得
Set cFileList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct( sBakDstDirPath, cFileList, 1, "*")

'バックアップファイル数取得＋既存の最古ファイル探索
Dim lBakFileNum
Dim sDelFilePath
lBakFileNum = 0
sDelFilePath = ""
For Each sFilePath in cFileList
    If ( (InStr(sFilePath, sBakDstPathBase) > 0) And _
         (objFSO.GetExtensionName(sFilePath) = sBakSrcFileExt) ) Then
        If lBakFileNum = 0 Then
           sDelFilePath = sFilePath
        End If
        lBakFileNum = lBakFileNum + 1
    End If
Next

'バックアップファイル削除
If lBakFileNum > lBakFileNumMax Then
    objFSO.DeleteFile sDelFilePath, True
End If

'objLogFile.WriteLine "バックアップ完了！", vbOKOnly, sSCRIPT_NAME

objLogFile.Close

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
