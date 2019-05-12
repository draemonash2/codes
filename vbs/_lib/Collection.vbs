Option Explicit

' ==================================================================
' = �T�v    �e�L�X�g�t�@�C���̒��g��z��Ɋi�[
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [out]   �t�@�C���̒��g
' = �ߒl    �ǂݏo������    Boolean             �ǂݏo������
' =                                                 True:�t�@�C������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' ==================================================================
Public Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.add objTxtFile.ReadLine
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
    On Error Goto 0
End Function
'   Call Test_OpenTxtFile2Array()
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        sFilePath = "C:\codes\vbs\��������CSV���`�c�[��\data_type_list_.csv"
        Dim bRet
        bRet = ReadTxtFileToCollection( sFilePath, cFileList )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox bRet
        MsgBox sOutput
    End Sub

' ==================================================================
' = �T�v    �z��̒��g���e�L�X�g�t�@�C���ɏ����o��
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [in]    �t�@�C���̒��g
' = �ߒl    �����o������    Boolean             �����o������
' =                                                 True:�����o������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' ==================================================================
Public Function WriteTxtFileFrCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objTxtFile
    Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
    
    If Err.Number = 0 Then
        Dim sFileLine
        For Each sFileLine In cFileContents
            objTxtFile.WriteLine sFileLine
        Next
        WriteTxtFileFrCollection = True
    Else
        WriteTxtFileFrCollection = False
    '   WScript.Echo "�G���[ " & Err.Description
    End If
    
    objTxtFile.Close
    On Error Goto 0
End Function
'   Call Test_WriteTxtFileFrCollection()
    Private Sub Test_WriteTxtFileFrCollection()
        Dim cFileContents
        Set cFileContents = CreateObject("System.Collections.ArrayList")
        cFileContents.Add "a"
        cFileContents.Add "b"
        cFileContents.Insert 1, "c"
        DIm sTrgtFilePath
        sTrgtFilePath = "C:\codes\vbs\��������CSV���`�c�[��\Test.csv"
        call WriteTxtFileFrCollection( sTrgtFilePath, cFileContents )
    End Sub
