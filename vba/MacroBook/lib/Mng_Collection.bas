Attribute VB_Name = "Mng_Collection"
Option Explicit

' collection manage library v1.01

' ==================================================================
' = 概要    テキストファイルの中身を配列に格納
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [out]   ファイルの中身
' = 戻値    読み出し結果    Boolean             読み出し結果
' =                                                 True:ファイル存在
' =                                                 False:それ以外
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Collection.bas
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
        '   WScript.Echo "エラー " & Err.Description
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
' = 概要    配列の中身をテキストファイルに書き出し
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [in]    ファイルの中身
' = 引数    bOverwrite      Boolean     [in]    True:上書き、False:新規ファイル
' = 戻値    書き出し結果    Boolean             書き出し結果
' =                                                 True:書き出し成功
' =                                                 False:それ以外
' = 覚書    なし
' = 依存    Mng_FileSys.bas/GetFileNotExistPath()
' = 所属    Mng_Collection.bas
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
    '   WScript.Echo "エラー " & Err.Description
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

