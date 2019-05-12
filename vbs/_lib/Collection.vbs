Option Explicit

' ==================================================================
' = 概要    テキストファイルの中身を配列に格納
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [out]   ファイルの中身
' = 戻値    読み出し結果    Boolean             読み出し結果
' =                                                 True:ファイル存在
' =                                                 False:それ以外
' = 覚書    なし
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
        '   WScript.Echo "エラー " & Err.Description
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
        sFilePath = "C:\codes\vbs\試験結果CSV整形ツール\data_type_list_.csv"
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
' = 概要    配列の中身をテキストファイルに書き出し
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [in]    ファイルの中身
' = 戻値    書き出し結果    Boolean             書き出し結果
' =                                                 True:書き出し成功
' =                                                 False:それ以外
' = 覚書    なし
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
    '   WScript.Echo "エラー " & Err.Description
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
        sTrgtFilePath = "C:\codes\vbs\試験結果CSV整形ツール\Test.csv"
        call WriteTxtFileFrCollection( sTrgtFilePath, cFileContents )
    End Sub
