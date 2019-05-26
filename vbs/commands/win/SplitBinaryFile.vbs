'TODO:要コマンド化

'ファイルを分割する
'分割するバイトサイズは指定できる
'バイナリレベルで分割する

Option Explicit

Dim SPLIT_BYTE_SIZE
Dim asWSArgs
Dim oSrcStream
Dim oDstStream

'★★★ 分割ファイルサイズはここで変更 ★★★
SPLIT_BYTE_SIZE = 1024000

Set asWSArgs = WScript.Arguments

If asWSArgs.Count <> 1 Then
    WScript.Echo "Drag and drop only one file."
    WScript.Quit
End If

Dim sInFileName
Dim sOutFileName
Dim oFileSys
Dim bIsFileExist
Dim oFile

sInFileName = asWSArgs(0)

Set asWSArgs = Nothing

Set oFileSys = CreateObject("Scripting.FileSystemObject")
bIsFileExist = oFileSys.FileExists(sInFileName)
If bIsFileExist Then
    Set oFile = oFileSys.GetFile(sInFileName)
    If oFile.Size < SPLIT_BYTE_SIZE Then
        WScript.Echo "File size is too small. It requires more than " & FormatNumber(SPLIT_BYTE_SIZE, 0) & " byte." & vbCrLf & sInFileName & vbCrLf & "(" & FormatNumber(oFile.Size, 0) & " byte)"
        WScript.Quit
    End If
    WScript.Echo "A target file is " & sInFileName & vbCrLf & "(" & FormatNumber(oFile.Size, 0) & " byte)"
    Set oFile = Nothing
Else
    WScript.Echo "No file."
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
    oDstStream.Write oSrcStream.Read( SPLIT_BYTE_SIZE )
    oDstStream.SaveToFile sOutFileName, 2
    oDstStream.Close
    oDstStream.Open
    lFileNum = lFileNum + 1
Loop

WScript.Echo sInFileName & " -> Success : " & lFileNum

oSrcStream.Close
oDstStream.Close

Set oSrcStream = Nothing
Set oDstStream = Nothing
