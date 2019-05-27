'ファイルを分割する
'分割するバイトサイズは指定できる
'バイナリレベルで分割する
'
'usage
' cscript.exe .\SplitBinaryFile.vbs <file_path> [<split_byte_size>]

Option Explicit

'==========================================================
'= 設定
'==========================================================
Const SPLIT_BYTE_SIZE = 1024000

'==========================================================
'= 本処理
'==========================================================
Dim oSrcStream
Dim oDstStream
Dim lSplitByteSize
Dim sInFileName
lSplitByteSize = SPLIT_BYTE_SIZE
If WScript.Arguments.Count = 2 Then
	sInFileName = WScript.Arguments(0)
	lSplitByteSize = WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 1 Then
	sInFileName = WScript.Arguments(0)
Else
    WScript.StdOut.WriteLine "argument error"
    WScript.Quit
End If

Dim sOutFileName
Dim oFileSys
Dim bIsFileExist
Dim oFile

Set oFileSys = CreateObject("Scripting.FileSystemObject")
bIsFileExist = oFileSys.FileExists(sInFileName)
If bIsFileExist Then
    Set oFile = oFileSys.GetFile(sInFileName)
    If oFile.Size < lSplitByteSize Then
        WScript.StdOut.WriteLine "File size is too small. It requires more than " & FormatNumber(lSplitByteSize, 0) & " byte." & vbCrLf & sInFileName & vbCrLf & "(" & FormatNumber(oFile.Size, 0) & " byte)"
        WScript.Quit
    End If
    WScript.StdOut.WriteLine "A target file is " & sInFileName & vbCrLf & "(" & FormatNumber(oFile.Size, 0) & " byte)"
    Set oFile = Nothing
Else
    WScript.StdOut.WriteLine "No file."
    WScript.Quit
End If

Set oFileSys = Nothing

Set oSrcStream = CreateObject("ADODB.Stream")
oSrcStream.Type = 1
oSrcStream.Open
oSrcStream.LoadFromFile sInFileName
oSrcStream.Position = 0

Set oDstStream = CreateObject("ADODB.Stream")
oDstStream.Type = 1
oDstStream.Open
oDstStream.Position = 0

Dim lFileNum
lFileNum = 0
Do While oSrcStream.EOS = False
    sOutFileName = sInFileName & "." & lFileNum
    oDstStream.Write oSrcStream.Read( lSplitByteSize )
    oDstStream.SaveToFile sOutFileName, 2
    oDstStream.Close
    oDstStream.Open
    lFileNum = lFileNum + 1
Loop

WScript.StdOut.WriteLine sInFileName & " -> Success : " & lFileNum

oSrcStream.Close
oDstStream.Close

Set oSrcStream = Nothing
Set oDstStream = Nothing
