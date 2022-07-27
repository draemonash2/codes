Attribute VB_Name = "Mng_Sys"
Option Explicit

Public Enum E_PATH_TYPE
    PATH_TYPE_FILE
    PATH_TYPE_DIRECTORY
    PATH_TYPE_NOT_EXIST
End Enum

Public Type T_PATH_LIST
    sPath As String
    ePathType As E_PATH_TYPE
End Type

Public gatPathList() As T_PATH_LIST

Public Function SysInit()
    Dim atPathList() As T_PATH_LIST
    gatPathList = atPathList
End Function

'�t�@�C�����i�g���q����j�擾
Public Function GetFileName( _
    ByVal sFilePath As String _
) As String
    Dim asFileName() As String
    Debug.Assert sFilePath <> ""
    asFileName = Split(sFilePath, "\")
    GetFileName = asFileName(UBound(asFileName))
    Debug.Assert GetFileName <> ""
End Function

'�f�B���N�g�����擾
Public Function GetParentDirName( _
    ByVal sFilePath As String _
) As String
    Dim asFileName() As String
    Debug.Assert sFilePath <> ""
    Debug.Assert InStr(sFilePath, ".") > 0
    asFileName = Split(sFilePath, "\")
    GetParentDirName = asFileName(UBound(asFileName) - 1)
    Debug.Assert GetParentDirName <> ""
End Function

'�t�@�C�����i�g���q�Ȃ��j�擾
Public Function GetFileNameBase( _
    ByVal sFilePath As String _
) As String
    Debug.Assert sFilePath <> ""
    GetFileNameBase = Split(GetFileName(sFilePath), ".")(0)
    Debug.Assert GetFileNameBase <> ""
End Function

'�g���q�擾
Public Function GetFileNameExt( _
    ByVal sFilePath As String _
) As String
    Debug.Assert sFilePath <> ""
    Debug.Assert InStr(sFilePath, ".")
    GetFileNameExt = Split(GetFileName(sFilePath), ".")(1)
    Debug.Assert GetFileNameExt <> ""
End Function

'�f�B���N�g���p�X�擾
Public Function GetDirPath( _
    ByVal sFilePath As String _
) As String
    Dim sFileName As String
    Debug.Assert sFilePath <> ""
    Debug.Assert Right(sFilePath, 1) <> "\" '������"\"���f�B���N�g���p�X �̏ꍇ�G���[
    Debug.Assert InStr(sFilePath, "\") > 0 '�p�X�łȂ��ꍇ�G���[
    sFileName = GetFileName(sFilePath)
    GetDirPath = Left(sFilePath, Len(sFilePath) - Len("\" & sFileName))
    Debug.Assert GetDirPath <> ""
End Function

Public Function GetFileList( _
    ByVal sTargetDir As String _
)
    Call FolderSearch(sTargetDir, 0, gatPathList)
End Function
 
'atPathList() �Ƀt�@�C�����X�g���i�[�����B
'iPathNum �� �Ăяo���� 0 �Œ�Ƃ��邱�ƁB
Private Function FolderSearch( _
    ByVal sTargetDir As String, _
    ByVal iPathNum As Integer, _
    ByRef atPathList() As T_PATH_LIST _
) As Integer
    Dim oFolder As Object
    Dim oSubFolder As Object
    Dim oFile As Object
 
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(sTargetDir)
    
    '*** �t�H���_�\�� ***
    ReDim Preserve atPathList(iPathNum)
    atPathList(iPathNum).sPath = oFolder.Path
    atPathList(iPathNum).ePathType = PATH_TYPE_DIRECTORY
    iPathNum = iPathNum + 1
 
    '�t�H���_���̃T�u�t�H���_���
    '�i�T�u�t�H���_���Ȃ���΃��[�v���͒ʂ�Ȃ��j
    For Each oSubFolder In oFolder.SubFolders
        iPathNum = FolderSearch(oSubFolder.Path, iPathNum, atPathList) '�ċA�I�Ăяo��
    Next oSubFolder
 
    '*** �t�@�C���� ***
    For Each oFile In oFolder.Files
        ReDim Preserve atPathList(iPathNum)
        atPathList(iPathNum).sPath = oFile.Path
        atPathList(iPathNum).ePathType = PATH_TYPE_FILE
        iPathNum = iPathNum + 1
    Next oFile
    
    FolderSearch = iPathNum
End Function

Public Function AddBak2FilePath( _
    ByVal sFilePath As String _
) As String
    Dim sRetFilePath As String
    AddBak2FilePath = GetDirPath(sFilePath) & "\" & _
                      GetFileNameBase(sFilePath) & "_bak" & _
                      "." & GetFileNameExt(sFilePath)
End Function

Public Function AddSeqNo2FilePath( _
    ByVal sFilePath As String _
) As String
    Dim sRetFilePath As String
    Dim lFileIdx As Long
    sRetFilePath = sFilePath
    lFileIdx = 1
    Do
        sRetFilePath = GetDirPath(sFilePath) & "\" & _
                       GetFileNameBase(sFilePath) & "_" & Format(lFileIdx, "000") & _
                       "." & GetFileNameExt(sFilePath)
        lFileIdx = lFileIdx + 1
    Loop While ChkFileExist(sRetFilePath) = True '�t�@�C�����݊m�F
    AddSeqNo2FilePath = sRetFilePath
End Function

Public Function GetTypeFileOrFolder( _
    ByVal sChkTrgtPath As String _
) As E_PATH_TYPE
    Dim oFileSys As Object
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    If oFileSys.FolderExists(sChkTrgtPath) = True And _
       oFileSys.FileExists(sChkTrgtPath) = False Then
        GetTypeFileOrFolder = PATH_TYPE_DIRECTORY
    Else
        If oFileSys.FolderExists(sChkTrgtPath) = False And _
           oFileSys.FileExists(sChkTrgtPath) = True Then
            GetTypeFileOrFolder = PATH_TYPE_FILE
        Else
            GetTypeFileOrFolder = PATH_TYPE_NOT_EXIST
        End If
    End If
    Set oFileSys = Nothing
End Function

Public Function ChkFileExist( _
    ByVal sFilePath As String _
) As Boolean
    Dim oFileSysObj As Object
    Debug.Assert sFilePath <> ""
    'Dir(sFilePath) �ł��t�@�C���̑��݊m�F�͂ł��邪�A
    '�����������̂��߃G���[����������ꍇ�����邽�߁A�g�p���Ȃ�
    Set oFileSysObj = CreateObject("Scripting.FileSystemObject")
    If oFileSysObj.FileExists(sFilePath) Then
        ChkFileExist = True
    Else
        ChkFileExist = False
    End If
    Set oFileSysObj = Nothing
End Function

Public Function CreBackupFile( _
    ByVal sFilePath As String _
)
    Dim oFileSysObj As Object
    Dim sSrcFilePath As String
    Dim sDstFilePath As String
    
    Set oFileSysObj = CreateObject("Scripting.FileSystemObject")
    sSrcFilePath = sFilePath
    sDstFilePath = AddBak2FilePath(sFilePath)
    oFileSysObj.CopyFile sSrcFilePath, sDstFilePath, True
    Set oFileSysObj = Nothing
End Function

'�������e�X�g�p������
Sub test()
    Dim asFilePath() As String
    Dim lProcIdx As Long
    Dim lProcNum As Long
    
    lProcNum = 0
    ReDim Preserve asFilePath(lProcNum): asFilePath(lProcNum) = "C:\Coffer\16686.xls": lProcNum = lProcNum + 1:
    ReDim Preserve asFilePath(lProcNum): asFilePath(lProcNum) = "C:\Coffer\": lProcNum = lProcNum + 1
    ReDim Preserve asFilePath(lProcNum): asFilePath(lProcNum) = "C:\Coffer": lProcNum = lProcNum + 1
    ReDim Preserve asFilePath(lProcNum): asFilePath(lProcNum) = "16686.xls": lProcNum = lProcNum + 1
    ReDim Preserve asFilePath(lProcNum): asFilePath(lProcNum) = "": lProcNum = lProcNum + 1
    
    For lProcIdx = 0 To lProcNum - 1
        Debug.Print Format(lProcIdx, "00") & ":" & GetDirPath(asFilePath(lProcIdx))
    Next lProcIdx
    Debug.Print ""
End Sub

Sub test2()
    Call CreBackupFile("C:\Coffer\svn_PTM\trunk\03_�Ǘ�\30_�c�[��\test\TF2���v���g2nd�P�̎������ڏ�_02_�f�f���j�^�ύX_dgMonCanLocalComErr-16686.xls")
End Sub


