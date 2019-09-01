Attribute VB_Name = "Mng_Collection"
Option Explicit

' collection manage library v1.01

' ==================================================================
' = �T�v    �e�L�X�g�t�@�C���̒��g��z��Ɋi�[
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [out]   �t�@�C���̒��g
' = �ߒl    �ǂݏo������    Boolean             �ǂݏo������
' =                                                 True:�t�@�C������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Collection.bas
' ==================================================================
Public Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath As String, _
    ByRef cFileContents As Collection _
)
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile As Object
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.Add objTxtFile.ReadLine
            Loop
            ReadTxtFileToCollection = True
        Else
            ReadTxtFileToCollection = False
        '   WScript.Echo "�G���[ " & Err.Description
        End If
        
        objTxtFile.Close
    Else
        ReadTxtFileToCollection = False
    End If
    On Error GoTo 0
End Function
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileContents As Collection
        Set cFileContents = New Collection
        Dim sInFilePath As String
        sInFilePath = "C:\codes\vbs\_lib\Test.csv"
        Dim bRet As Boolean
        bRet = ReadTxtFileToCollection(sInFilePath, cFileContents)
    End Sub

' ==================================================================
' = �T�v    �z��̒��g���e�L�X�g�t�@�C���ɏ����o��
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [in]    �t�@�C���̒��g
' = ����    bOverwrite      Boolean     [in]    True:�㏑���AFalse:�V�K�t�@�C��
' = �ߒl    �����o������    Boolean             �����o������
' =                                                 True:�����o������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/GetFileNotExistPath()
' = ����    Mng_Collection.bas
' ==================================================================
Public Function WriteTxtFileFrCollection( _
    ByVal sTrgtFilePath As String, _
    ByRef cFileContents As Collection, _
    ByVal bOverwrite As Boolean _
)
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objTxtFile As Object
    If bOverwrite = True Then
        'Do Nothing
    Else
        Dim sInTrgtFilePath
        sInTrgtFilePath = sTrgtFilePath
        sTrgtFilePath = GetFileNotExistPath(sInTrgtFilePath)
    End If
    Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
    
    If Err.Number = 0 Then
        Dim vFileLine As Variant
        For Each vFileLine In cFileContents
            objTxtFile.WriteLine vFileLine
        Next
        WriteTxtFileFrCollection = True
    Else
        WriteTxtFileFrCollection = False
    '   WScript.Echo "�G���[ " & Err.Description
    End If
    
    objTxtFile.Close
    On Error GoTo 0
End Function
    Private Sub Test_WriteTxtFileFrCollection()
        Dim cFileContents As Collection
        Set cFileContents = New Collection
        cFileContents.Add "a"
        cFileContents.Add "fff"
        cFileContents.Add "d"
        cFileContents.Add "e"
        cFileContents.Add Item:="c", after:=1
        Dim sTrgtFilePath As String
        sTrgtFilePath = "C:\codes\vbs\_lib\Test.csv"
        Call WriteTxtFileFrCollection(sTrgtFilePath, cFileContents, False)
    End Sub

