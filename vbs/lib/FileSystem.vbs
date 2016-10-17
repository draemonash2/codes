Option Explicit

'lFileListType�j0�F�����A1:�t�@�C���A2:�t�H���_�A����ȊO�F�i�[���Ȃ�
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
    
    '*** �t�H���_�p�X�i�[ ***
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
    
    '�t�H���_���̃T�u�t�H���_���
    '�i�T�u�t�H���_���Ȃ���΃��[�v���͒ʂ�Ȃ��j
    For Each objSubFolder In objFolder.SubFolders
        Call GetFileList( objSubFolder, asFileList, lFileListType)
    Next
    
    '*** �t�@�C���p�X�i�[ ***
    For Each objFile In objFolder.Files
        Select Case lFileListType
            Case 0:    bExecStore = True
            Case 1:    bExecStore = True
            Case 2:    bExecStore = False
            Case Else: bExecStore = False
        End Select
        If bExecStore = True Then
            '�{�X�N���v�g�t�@�C���͊i�[�ΏۊO
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

'�t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
Public Function CreateDirectry( _
    ByVal sDirPath _
)
    Dim sParentDir
    Dim oFileSys
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
    
    '�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
    If oFileSys.FolderExists( sParentDir ) = False Then
        Call CreateDirectry( sParentDir )
    End If
    
    '�f�B���N�g���쐬
    If oFileSys.FolderExists( sDirPath ) = False Then
        oFileSys.CreateFolder sDirPath
    End If
    
    Set oFileSys = Nothing
End Function

'�߂�l�j1�F�t�@�C���A2�A�t�H���_�[�A0�F�G���[�i���݂��Ȃ��p�X�j
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
        GetFileOrFolder = 1 '�t�@�C��
    ElseIf bFolderExists = True And bFileExists = False Then
        GetFileOrFolder = 2 '�t�H���_�[
    Else
        GetFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
    End If
End Function
