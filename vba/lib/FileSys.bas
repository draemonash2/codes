Attribute VB_Name = "FileSys"
Option Explicit

' file system library v1.0

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
' ============================================
' = 概要    ファイルの内容を配列に読み込む。
' = 引数    sFilePath   String   入力するファイルパス
' =         sCharSet    String   キャラクタセット
' = 戻値                String() ファイル内容
' = 覚書    なし
' ============================================
Public Function InputTxtFile( _
    ByRef sFilePath As String, _
    Optional ByVal sCharSet As String _
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

' ============================================
' = 概要    配列の内容をファイルに書き込む。
' = 引数    sFilePath     String  [in]  出力するファイルパス
' =         asFileLine()  String  [in]  出力するファイルの内容
' = 戻値    なし
' = 覚書    なし
' ============================================
Public Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String _
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

'フォルダが既に存在している場合は何もしない
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

'atPathList() にファイルリストが格納される。
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

