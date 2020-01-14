Attribute VB_Name = "Mng_FileSys"
Option Explicit

' file system library v1.5

Public Enum E_PATH_TYPE
    PATH_TYPE_FILE
    PATH_TYPE_DIRECTORY
End Enum
 
Public Type T_PATH_LIST
    sPath As String
    sName As String
    ePathType As E_PATH_TYPE
End Type
 
Public Enum T_SYSOBJ_TYPE
    SYSOBJ_NOT_EXIST
    SYSOBJ_FILE
    SYSOBJ_DIRECTORY
End Enum
  
'参照設定「Microsoft ActiveX Data Objects 6.1 Liblary」をチェックすること！
' ==================================================================
' = 概要    ファイルの内容を配列に読み込む。
' = 引数    sFilePath   String   入力するファイルパス
' =         sCharSet    String   キャラクタセット
' = 戻値                String() ファイル内容
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function InputTxtFile( _
    ByRef sFilePath As String, _
    Optional ByVal sCharSet As String = "shift_jis" _
) As String()
    Dim lLineCnt As Long: lLineCnt = 0
    Dim asRetStr() As String
    Dim oTxtObj As Object
    
    Set oTxtObj = CreateObject("ADODB.Stream")
    
    With oTxtObj
        .Type = adTypeText           'オブジェクトに保存するデータの種類を文字列型に指定する
        .Charset = sCharSet
        .Open
        .LoadFromFile (sFilePath)
        
        lLineCnt = 0
        Do While Not .EOS
            ReDim Preserve asRetStr(lLineCnt)
            asRetStr(lLineCnt) = .ReadText(adReadLine)
            lLineCnt = lLineCnt + 1
        Loop
        
        .Close
    End With
    
    Set oTxtObj = Nothing
    
    InputTxtFile = asRetStr
    
End Function

' ==================================================================
' = 概要    配列の内容をファイルに書き込む。
' = 引数    sFilePath     String  [in]  出力するファイルパス
' =         asFileLine()  String  [in]  出力するファイルの内容
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String = "shift_jis" _
)
    Dim oTxtObj As Object
    Dim lLineIdx As Long
    
    If Sgn(asFileLine) = 0 Then
        'Do Nothing
    Else
        Set oTxtObj = CreateObject("ADODB.Stream")
        With oTxtObj
            .Type = adTypeText
            .Charset = sCharSet
            .Open
            
            '配列を1行ずつオブジェクトに書き込む
            For lLineIdx = 0 To UBound(asFileLine)
                .WriteText asFileLine(lLineIdx), adWriteLine
            Next lLineIdx
            
            .SaveToFile (sFilePath), adSaveCreateOverWrite    'オブジェクトの内容をファイルに保存
            .Close
        End With
    End If
    
    Set oTxtObj = Nothing
End Function

' ==================================================================
' = 概要    ディレクトリを作成する。親ディレクトリも自動生成する。
' = 引数    sDirPath    String  [in]  フォルダパス
' = 戻値    なし
' = 覚書    フォルダが既に存在している場合は何もしない
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
 
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
 
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
 
    '親ディレクトリが存在しない場合、再帰呼び出し
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
 
    'ディレクトリ作成
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
 
    Set oFileSys = Nothing
End Function

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する
' = 引数    sTrgtDir        String      [in]    対象フォルダ
' = 引数    atPathList      T_PATH_LIST [out]   ファイル/フォルダパス一覧
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetFileList( _
    ByVal sTargetDir As String, _
    ByRef atPathList() As T_PATH_LIST _
)
    Dim oFolder As Object
    Dim oSubFolder As Object
    Dim oFile As Object
    Dim lLastIdx As Long
 
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(sTargetDir)
 
    '*** フォルダ列挙 ***
    If Sgn(atPathList) = 0 Then
        ReDim Preserve atPathList(0)
    Else
        ReDim Preserve atPathList(UBound(atPathList) + 1)
    End If
    lLastIdx = UBound(atPathList)
    atPathList(lLastIdx).sPath = oFolder.Path
    atPathList(lLastIdx).sName = oFolder.Name
    atPathList(lLastIdx).ePathType = PATH_TYPE_DIRECTORY
 
    'フォルダ内のサブフォルダを列挙
    '（サブフォルダがなければループ内は通らない）
    For Each oSubFolder In oFolder.SubFolders
        Call GetFileList(oSubFolder.Path, atPathList) '再帰的呼び出し
    Next oSubFolder
 
    '*** ファイル列挙 ***
    For Each oFile In oFolder.Files
        If Sgn(atPathList) = 0 Then
            ReDim Preserve atPathList(0)
        Else
            ReDim Preserve atPathList(UBound(atPathList) + 1)
        End If
        lLastIdx = UBound(atPathList)
        atPathList(lLastIdx).sPath = oFile.Path
        atPathList(lLastIdx).sName = oFile.Name
        atPathList(lLastIdx).ePathType = PATH_TYPE_FILE
    Next oFile
End Function

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する(Variant,Dirコマンド版)
' = 引数    sTrgtDir        String      [in]    対象フォルダ
' = 引数    vFileList       Variant     [out]   ファイル/フォルダパス一覧
' = 引数    lFileListType   Long        [in]    取得する一覧の形式
' =                                                 0：両方
' =                                                 1:ファイル
' =                                                 2:フォルダ
' =                                                 それ以外：格納しない
' = 戻値    なし
' = 覚書    ・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' =         ・vFileList は配列型ではなくバリアント型として定義する
' =           必要があることに注意！
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetFileList2( _
    ByVal sTrgtDir As String, _
    ByRef vFileList As Variant, _
    ByVal lFileListType As Long _
)
    Dim objFSO As Object 'FileSystemObjectの格納先
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dir コマンド実行（出力結果を一時ファイルに格納）
    Dim sTmpFilePath As String
    Dim sExecCmd As String
    sTmpFilePath = CreateObject("WScript.Shell").CurrentDirectory & "\Dir.tmp"
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile As Object
    Dim sTextAll As String
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(sTmpFilePath, 1)
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left(sTextAll, Len(sTextAll) - Len(vbNewLine))       '末尾に改行が付与されてしまうため、削除
            vFileList = Split(sTextAll, vbNewLine)
            objFile.Close
        Else
            MsgBox "ファイルが開けません: " & Err.Description
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        MsgBox "エラー " & Err.Description
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error GoTo 0
End Function
    Private Sub Test_GetFileList2()
        Dim objWshShell As Object
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sCurDir As String
        sCurDir = "C:\codes"
        
        Dim vFileList As Variant
'        Call GetFileList2("C:\codes", vFileList, 0)
'        Call GetFileList2("C:\codes", vFileList, 1)
        Call GetFileList2("C:\codes", vFileList, 2)
    End Sub

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する(Collection,Dirコマンド版)
' = 引数    sTrgtDir        String              [in]    対象フォルダ
' = 引数    cFileList       Object(Collection)  [out]   ファイル/フォルダパス一覧
' = 引数    lFileListType   Long                [in]    取得する一覧の形式
' =                                                         0：両方
' =                                                         1:ファイル
' =                                                         2:フォルダ
' =                                                         それ以外：格納しない
' = 引数    sFileExtStr     String              [in]    取得するファイルの拡張子(省略可能)
' =                                                       ex1) ""
' =                                                       ex2) "*"
' =                                                       ex3) "*.c"
' =                                                       ex4) "*.txt *.log *.csv"
' = 戻値    なし
' = 覚書    ・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' = 覚書    ・sFileExtStrはファイル指定時のみ有効
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetObjctListCmdClct( _
    ByVal sTrgtDir As String, _
    ByRef cFileList As Object, _
    ByVal lFileListType As Long, _
    Optional ByVal sFileExtStr As String = "" _
)
    Dim objFSO As Object 'FileSystemObjectの格納先
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dir コマンド実行（出力結果を一時ファイルに格納）
    Dim sTmpFilePath As String
    Dim sExecCmd As String
    sTmpFilePath = CreateObject("WScript.Shell").CurrentDirectory & "\Dir.tmp"
    Dim sTrgtDirStr As String
    If sFileExtStr = "" Then
        sTrgtDirStr = """" & sTrgtDir & """"
    Else
        Dim vFileExtentions As Variant
        vFileExtentions = Split(sFileExtStr, " ")
        Dim lSplitIdx As Long
        For lSplitIdx = 0 To UBound(vFileExtentions)
            If sTrgtDirStr = "" Then
                sTrgtDirStr = """" & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            Else
                sTrgtDirStr = sTrgtDirStr & " """ & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            End If
        Next lSplitIdx
    End If
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile As Object
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(sTmpFilePath, 1)
        If Err.Number = 0 Then
            Do Until objFile.AtEndOfStream
                cFileList.Add objFile.ReadLine
            Loop
        Else
            MsgBox "ファイルが開けません: " & Err.Description
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        MsgBox "エラー " & Err.Description
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error GoTo 0
End Function
    Private Sub Test_GetObjctListCmdClct()
        Dim sRootDir As String
        sRootDir = "C:\codes"
        
        Dim cFileList As Object
        Set cFileList = CreateObject("System.Collections.ArrayList")
        
'        Call GetObjctListCmdClct(sRootDir, cFileList, 0)
        Call GetObjctListCmdClct(sRootDir, cFileList, 1)
'        Call GetObjctListCmdClct(sRootDir, cFileList, 1, "*.c *.h")
'        Call GetObjctListCmdClct(sRootDir, cFileList, 1, "*.vbs")
'        Call GetObjctListCmdClct(sRootDir, cFileList, 1, "*")
'        Call GetObjctListCmdClct(sRootDir, cFileList, 1, "")
'        Call GetObjctListCmdClct(sRootDir, cFileList, 2)
        Stop
    End Sub

' ==================================================================
' = 概要    ファイル/フォルダを判定する
' = 引数    sChkTrgtPath    String          [in]    対象ファイルパス
' = 戻値                    T_SYSOBJ_TYPE           ファイルorフォルダ
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetFileOrFolder( _
    ByVal sChkTrgtPath As String _
) As T_SYSOBJ_TYPE
    Dim oFileSys As Object
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    If oFileSys.FolderExists(sChkTrgtPath) = True And _
       oFileSys.FileExists(sChkTrgtPath) = False Then
        GetFileOrFolder = SYSOBJ_DIRECTORY
    Else
        If oFileSys.FolderExists(sChkTrgtPath) = False And _
           oFileSys.FileExists(sChkTrgtPath) = True Then
            GetFileOrFolder = SYSOBJ_FILE
        Else
            GetFileOrFolder = SYSOBJ_NOT_EXIST
        End If
    End If
    Set oFileSys = Nothing
End Function

' ==================================================================
' = 概要    フォルダ選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトフォルダパス（省略可）
' = 戻値                String        フォルダ選択結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFolderSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists(sSelectedPath) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        MsgBox ShowFolderSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") _
                )
    End Sub

' ==================================================================
' = 概要    ファイル（単一）選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトファイルパス（省略可）
' = 引数    sFilters    String  [in]  選択時のフィルタ（省略可）(※)
' = 戻値                String        ファイル選択結果
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    ・拡張子が複数ある場合は、";"で区切る
' =                    ・ファイル種別と拡張子は"/"で区切る
' =                    ・フィルタが複数ある場合、","で区切る
' =         sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' = 依存    Mng_FileSys.bas/SetDialogFilters()
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイルを選択してください"
    fdDialog.AllowMultiSelect = False
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
 
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFileSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
            ShowFileSelectDialog = sSelectedPath
        Else
            ShowFileSelectDialog = ""
        End If
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFileSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        sFilters = ""
        
        MsgBox ShowFileSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
    '    MsgBox ShowFileSelectDialog( _
    '                objWshShell.SpecialFolders("Desktop") & "\test.txt" _
    '            )
    End Sub

' ==================================================================
' = 概要    ファイル（複数）選択ダイアログを表示する
' = 引数    asSelectedFiles String()    [out] 選択されたファイルパス一覧
' = 引数    sInitPath       String      [in]  デフォルトファイルパス（省略可）
' = 引数    sFilters        String      [in]  選択時のフィルタ（省略可）(※)
' = 戻値    なし
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    ・拡張子が複数ある場合は、";"で区切る
' =                    ・ファイル種別と拡張子は"/"で区切る
' =                    ・フィルタが複数ある場合、","で区切る
' =         sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' = 依存    Mng_FileSys.bas/SetDialogFilters()
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFilesSelectDialog( _
    ByRef asSelectedFiles() As String, _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
)
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイルを選択してください（複数可）"
    fdDialog.AllowMultiSelect = True
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
 
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ReDim Preserve asSelectedFiles(0)
        asSelectedFiles(0) = ""
    Else
        Dim lSelNum As Long
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asSelectedFiles(lSelNum - 1)
        Dim lSelIdx As Long
        For lSelIdx = 0 To lSelNum - 1
            Dim sSelectedPath As String
            sSelectedPath = fdDialog.SelectedItems(lSelIdx + 1)
            If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
                asSelectedFiles(lSelIdx) = sSelectedPath
            Else
                asSelectedFiles(lSelIdx) = ""
            End If
        Next lSelIdx
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFilesSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        sFilters = "全てのファイル/*.*,画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
 
        Dim asSelectedFiles() As String
        Call ShowFilesSelectDialog( _
                    asSelectedFiles, _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
        Dim sBuf As String
        sBuf = ""
        sBuf = sBuf & vbNewLine & UBound(asSelectedFiles) + 1
        Dim lSelIdx As Long
        For lSelIdx = 0 To UBound(asSelectedFiles)
            sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
        Next lSelIdx
        MsgBox sBuf
    End Sub

' ==================================================================
' = 概要    ShowFileSelectDialog() と ShowFilesSelectDialog() 用の関数
' =         ダイアログのフィルタを追加する。指定方法は以下。
' =           ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =               ・拡張子が複数ある場合は、";"で区切る
' =               ・ファイル種別と拡張子は"/"で区切る
' =               ・フィルタが複数ある場合、","で区切る
' = 引数    sFilters    String      [in]    フィルタ
' = 引数    fdDialog    FileDialog  [in]    ファイルダイアログ
' = 戻値    なし
' = 覚書    sFilters が空文字の場合、フィルタをクリアする。
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                            """/"" は一つだけ指定してください" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "処理を中断します。"
                        End
                    End If
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "処理を中断します。"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        """/"" は一つだけ指定してください" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "処理を中断します。"
                    End
                End If
            Else
                MsgBox _
                    "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                    "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                    "  " & sFilters
                MsgBox "処理を中断します。"
                End
            End If
        End If
    End If
End Function

' ==================================================================
' = 概要    指定パスが存在する場合、"_XXX" を付与して返却する
' = 引数    sTrgtPath       String      [in]    対象パス
' = 引数    sAddedPath      String      [out]   付与後のパス
' = 引数    lAddedPathType  Long        [out]   付与後のパス種別
' =                                               1: ファイル
' =                                               2: フォルダ
' = 戻値                    Boolean             取得結果
' = 覚書    本関数では、ファイル/フォルダは作成しない。
' = 依存    Mng_FileSys.bas/GetFileNotExistPath()
' =         Mng_FileSys.bas/GetFolderNotExistPath()
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetNotExistPath( _
    ByVal sTrgtPath As String, _
    ByRef sAddedPath As String, _
    ByRef lAddedPathType As Long _
) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim bFolderExists As Boolean
    Dim bFileExists As Boolean
    bFolderExists = objFSO.FolderExists(sTrgtPath)
    bFileExists = objFSO.FileExists(sTrgtPath)
    
    If bFolderExists = False And bFileExists = True Then
        sAddedPath = GetFileNotExistPath(sTrgtPath)
        lAddedPathType = 1
        GetNotExistPath = True
    ElseIf bFolderExists = True And bFileExists = False Then
        sAddedPath = GetFolderNotExistPath(sTrgtPath)
        lAddedPathType = 2
        GetNotExistPath = True
    Else
        sAddedPath = sTrgtPath
        lAddedPathType = 0
        GetNotExistPath = False
    End If
End Function
    Private Sub Test_GetNotExistPath()
        Dim sOutStr As String
        Dim sAddedPath As String
        Dim lAddedPathType As Long
        Dim bRet As Boolean
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

' ==================================================================
' = 概要    指定ファイルパスが存在する場合、"_XXX" を付与して返却する
' = 引数    sTrgtPath       String      [in]    対象パス
' = 戻値                    String              付与後パス
' = 覚書    本関数では、ファイルは作成しない。
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetFileNotExistPath( _
    ByVal sTrgtPath As String _
) As String
    Dim lIdx As Long
    Dim objFSO As Object
    Dim sFileParDirPath As String
    Dim sFileBaseName As String
    Dim sFileExtName As String
    Dim sCreFilePath As String
    Dim bIsTrgtPathExists As Boolean
    
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreFilePath = sTrgtPath
    bIsTrgtPathExists = False
    Do While objFSO.FileExists(sCreFilePath)
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sFileParDirPath = objFSO.GetParentFolderName(sTrgtPath)
        sFileBaseName = objFSO.GetBaseName(sTrgtPath) & "_" & String(3 - Len(CStr(lIdx)), "0") & CStr(lIdx)
        sFileExtName = objFSO.GetExtensionName(sTrgtPath)
        If sFileExtName = "" Then
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName
        Else
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName & "." & sFileExtName
        End If
    Loop
    GetFileNotExistPath = sCreFilePath
End Function
    Private Sub Test_GetFileNotExistPath()
        Dim sOutStr As String
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

'*********************************************************************
'* ローカル関数定義
'*********************************************************************
' ==================================================================
' = 概要    指定フォルダパスが存在する場合、"_XXX" を付与して返却する
' = 引数    sTrgtPath       String      [in]    対象パス
' = 戻値                    String              付与後パス
' = 覚書    本関数では、フォルダは作成しない。
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function GetFolderNotExistPath( _
    ByVal sTrgtPath As String _
) As String
    Dim lIdx As Long
    Dim objFSO As Object
    Dim sCreDirPath  As String
    Dim bIsTrgtPathExists As Boolean
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreDirPath = sTrgtPath
    bIsTrgtPathExists = False
    Do While objFSO.FolderExists(sCreDirPath)
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sCreDirPath = sTrgtPath & "_" & String(3 - Len(CStr(lIdx)), "0") & CStr(lIdx)
    Loop
    If bIsTrgtPathExists = True Then
        GetFolderNotExistPath = sCreDirPath
    Else
        GetFolderNotExistPath = ""
    End If
End Function
    Private Sub Test_GetFolderNotExistPath()
        Dim sOutStr As String
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba")
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub
