Option Explicit

Private Function DownloadFile( _
    ByVal sDownloadUrl, _
    ByVal sLocalFilePath _
)
    Dim bResult
    bResult = True
    ' ダウンロード用のオブジェクト
    Dim objSrvHTTP
    Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")
    on error resume next
    Call objSrvHTTP.Open("GET", sDownloadUrl, False )
    if Err.Number <> 0 then
        Wscript.Echo Err.Description
        bResult = False
        Exit Function
    end if
    objSrvHTTP.Send
    
    if Err.Number <> 0 then
    ' おそらくサーバーの指定が間違っている
        Wscript.Echo Err.Description
        bResult = False
        Exit Function
    end if
    on error goto 0
    if objSrvHTTP.status = 404 then
        Wscript.Echo "URLが正しくありません(404)"
        bResult = False
        Exit Function
    end if
    
    ' バイナリデータ保存用オブジェクト
    Dim Stream
    Set Stream = Wscript.CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 ' バイナリ
    Stream.Write objSrvHTTP.responseBody    ' 戻されたバイナリをファイルとしてストリームに書き込み
    Stream.SaveToFile sLocalFilePath, 2     ' ファイルとして保存
    Stream.Close
    DownloadFile = bResult
End Function
'   Call Test_DownloadFile()
    Private Function Test_DownloadFile()
        Dim sDownloadTrgtFilePath
        sDownloadTrgtFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\codes-master.zip"
        Call DownloadFile("https://github.com/draemonash2/codes/archive/master.zip", sDownloadTrgtFilePath)
    End Function

