Option Explicit

Private Function DownloadFile( _
    ByVal sDownloadUrl, _
    ByVal sLocalFilePath _
)
    Dim bResult
    bResult = True
    ' �_�E�����[�h�p�̃I�u�W�F�N�g
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
    ' �����炭�T�[�o�[�̎w�肪�Ԉ���Ă���
        Wscript.Echo Err.Description
        bResult = False
        Exit Function
    end if
    on error goto 0
    if objSrvHTTP.status = 404 then
        Wscript.Echo "URL������������܂���(404)"
        bResult = False
        Exit Function
    end if
    
    ' �o�C�i���f�[�^�ۑ��p�I�u�W�F�N�g
    Dim Stream
    Set Stream = Wscript.CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 ' �o�C�i��
    Stream.Write objSrvHTTP.responseBody    ' �߂��ꂽ�o�C�i�����t�@�C���Ƃ��ăX�g���[���ɏ�������
    Stream.SaveToFile sLocalFilePath, 2     ' �t�@�C���Ƃ��ĕۑ�
    Stream.Close
    DownloadFile = bResult
End Function
'   Call Test_DownloadFile()
    Private Function Test_DownloadFile()
        Dim sDownloadTrgtFilePath
        sDownloadTrgtFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\codes-master.zip"
        Call DownloadFile("https://github.com/draemonash2/codes/archive/master.zip", sDownloadTrgtFilePath)
    End Function

