Option Explicit

'lFileListType）0：両方、1:ファイル、2:フォルダ、それ以外：格納しない
Public Function GetFileList( _
    ByVal sTrgtDir, _
    ByRef asFileList, _
    ByVal lFileListType _
)
    Dim objFileSys
    Dim objFolder
    Dim objSubFolder
    Dim objFile
    Dim bExecStore
    Dim lLastIdx
    
    Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSys.GetFolder( sTrgtDir )
    
    '*** フォルダパス格納 ***
    Select Case lFileListType
        Case 0:    bExecStore = True
        Case 1:    bExecStore = False
        Case 2:    bExecStore = True
        Case Else: bExecStore = False
    End Select
    If bExecStore = True Then
        lLastIdx = UBound( asFileList ) + 1
        ReDim Preserve asFileList( lLastIdx )
        asFileList( lLastIdx ) = objFolder
    Else
        'Do Nothing
    End If
    
    'フォルダ内のサブフォルダを列挙
    '（サブフォルダがなければループ内は通らない）
    For Each objSubFolder In objFolder.SubFolders
        Call GetFileList( objSubFolder, asFileList, lFileListType)
    Next
    
    '*** ファイルパス格納 ***
    For Each objFile In objFolder.Files
        Select Case lFileListType
            Case 0:    bExecStore = True
            Case 1:    bExecStore = True
            Case 2:    bExecStore = False
            Case Else: bExecStore = False
        End Select
        If bExecStore = True Then
            '本スクリプトファイルは格納対象外
            If objFile.Name = WScript.ScriptName Then
                'Do Nothing
            Else
                lLastIdx = UBound( asFileList ) + 1
                ReDim Preserve asFileList( lLastIdx )
                asFileList( lLastIdx ) = objFile
            End If
        Else
            'Do Nothing
        End If
    Next
    
    Set objFolder = Nothing
    Set objFileSys = Nothing
End Function

'フォルダが既に存在している場合は何もしない
Public Function CreateDirectry( _
    ByVal sDirPath _
)
    Dim sParentDir
    Dim oFileSys
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
    
    '親ディレクトリが存在しない場合、再帰呼び出し
    If oFileSys.FolderExists( sParentDir ) = False Then
        Call CreateDirectry( sParentDir )
    End If
    
    'ディレクトリ作成
    If oFileSys.FolderExists( sDirPath ) = False Then
        oFileSys.CreateFolder sDirPath
    End If
    
    Set oFileSys = Nothing
End Function

'戻り値）1：ファイル、2、フォルダー、0：エラー（存在しないパス）
Public Function GetFileOrFolder( _
    ByVal sChkTrgtPath _
)
    Dim oFileSys
    Dim bFolderExists
    Dim bFileExists
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    bFolderExists = oFileSys.FolderExists(sChkTrgtPath)
    bFileExists = oFileSys.FileExists(sChkTrgtPath)
    Set oFileSys = Nothing
    
    If bFolderExists = False And bFileExists = True Then
        GetFileOrFolder = 1 'ファイル
    ElseIf bFolderExists = True And bFileExists = False Then
        GetFileOrFolder = 2 'フォルダー
    Else
        GetFileOrFolder = 0 'エラー（存在しないパス）
    End If
End Function
