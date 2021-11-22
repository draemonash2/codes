Option Explicit

'ファイルを指定すると現在時刻を付与したバックアップファイルを作成する。
'同じファイル名のものが存在していたら、アルファベットを付与したバックアップファイルを作成する。
'   ex. 211201a, 211202b, …
'指定数分バックアップがたまったら、古いものから削除する。
'バックアップ対象はファイルのみ。
'第二引数にバックアップファイル数を指定できる。
'前回バックアップ時から更新されていない場合、バックアップしない。
'★要修正★バックアップ最大数分、同じ日付のバックアップファイルで満たされると、新しいファイルが更新されていくため要注意。

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
Const lMAX_BAK_FILE_NUM_DEFAULT = 50
Const sBAK_DIR_NAME = "_bak"
Const sBAK_FILE_SUFFIX = "#b#"

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "ファイルバックアップ"
Dim sBakSrcFilePath
Dim lBakFileNumMax
If WScript.Arguments.Count >= 2 Then
    sBakSrcFilePath = WScript.Arguments(0)
    lBakFileNumMax = CLng(WScript.Arguments(1))
ElseIf WScript.Arguments.Count = 1 Then
    sBakSrcFilePath = WScript.Arguments(0)
    lBakFileNumMax = lMAX_BAK_FILE_NUM_DEFAULT
Else
    WScript.Echo "引数を指定してください。プログラムを中断します。"
    WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

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

'既存の最新/最古ファイル探索
Dim lBakFileNum
Dim sBakDstFilePathOldest
Dim sBakDstFilePathLatest  '既存の最新バックアップファイル
lBakFileNum = 0
sBakDstFilePathOldest = ""
sBakDstFilePathLatest = ""
Dim vDateNewest
Dim vDateOldest
vDateOldest = Now()
Dim sFilePath
For Each sFilePath in cFileList
    If ( (InStr(sFilePath, sBakDstPathBase) > 0) And _
         (objFSO.GetExtensionName(sFilePath) = sBakSrcFileExt) ) Then
        Msgbox sFilePath
        Dim vDate
        Call GetFileInfo( sFilePath, 11, vDate)
        If vDate > vDateNewest Then
            vDateNewest = vDate
            sBakDstFilePathLatest = sFilePath
        End If
        If vDate < vDateOldest Then
            vDateOldest = vDate
            sBakDstFilePathOldest = sFilePath
        End If
        lBakFileNum = lBakFileNum + 1
    End If
Next
MsgBox "Old:" & sBakDstFilePathOldest & vbNewLine & _
       "New:" & sBakDstFilePathLatest
WScript.Quit

'バックアップファイル名確定
'既存のバックアップファイルが存在し、同じ日付のバックアップファイルが存在する場合
Dim sBakDstFilePath
If sBakDstFilePathLatest <> "" And _
   InStr(sBakDstFilePathLatest, sBakDstPathBase & sDateSuffix) > 0 Then
    Dim sTailChar
    sTailChar = Right( objFSO.GetBaseName( sBakDstFilePathLatest ), 1)
    Dim lBakDstAlphaIdx
    If Asc(sTailChar) >= Asc("a") And Asc(sTailChar) < Asc("z") Then
        lBakDstAlphaIdx = Asc(sTailChar) + 1
    ElseIf Asc(sTailChar) <= Asc("z") Then
        lBakDstAlphaIdx = Asc("a")
    ElseIf Asc(sTailChar) >= Asc("0") And Asc(sTailChar) <= Asc("9") Then
        lBakDstAlphaIdx = Asc("a")
    Else
        WScript.Echo "不正なバックアップファイルが見つかりました。"
        WScript.Echo "  " & sBakDstFilePathLatest
        WScript.Echo "プログラムを中断します。"
        WScript.Quit
    End If
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx) & "." & sBakSrcFileExt
Else
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & "." & sBakSrcFileExt
End If

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
    
    'バックアップファイル削除
    If lBakFileNum > lBakFileNumMax Then
        objFSO.DeleteFile sBakDstFilePathOldest, True
    End If
    
    'WScript.Echo "[Success] " & sBakSrcFilePath & " -> " & sBakDstFilePath
Else
    '前回バックアップ時から更新されていない場合、バックアップしない
    'WScript.Echo "[Skip]    " & sBakSrcFilePath
End If

'MsgBox "バックアップ完了！", vbOKOnly, sSCRIPT_NAME

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
