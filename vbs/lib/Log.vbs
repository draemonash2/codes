Option Explicit

Class LogMng
    Private gbLogFileEnable
    Private goLogFile
    
    Private Sub Class_Initialize()
        Call Init()
    End Sub
    Private Sub Class_Terminate()
        Call Close()
    End Sub
    
    Private Function Init()
        gbLogFileEnable = False
        Set goLogFile = Nothing
    End Function
    
    ' ==================================================================
    ' = �T�v    �o�̓��[�h��I������
    ' = ����    lSelectedMode   Long   [in]     �o�̓��[�h
    ' =                                           0:�W���o��
    ' =                                           1:���O�t�@�C���o��
    ' = �ߒl                    Boolean         �I������
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Function Mode( _
        ByVal lSelectedMode _
    )
        If lSelectedMode = 0 Then
            gbLogFileEnable = False
            Mode = True
        ElseIf lSelectedMode = 1 Then
            gbLogFileEnable = True
            Mode = True
        Else
            Mode = False
        End If
    End Function
    
    ' ==================================================================
    ' = �T�v    ���O�o�͂��J�n����
    ' = ����    sTrgtPath   String   [in]   �Ώۃt�@�C���p�X
    ' = ����    sIOMode     String   [in]   IO ���[�h
    ' =                                       "r":�Ǐo��
    ' =                                       "w":�V�K������
    ' =                                       "+w":�ǉ�������
    ' = �ߒl    �Ȃ�
    ' = �o��    �{�֐����Ăяo���ƁA�o�̓��[�h�������I�Ɂu���O�t�@�C��
    ' =         �o�́v�֐؂�ւ��
    ' ==================================================================
    Public Function Open( _
        ByVal sTrgtPath, _
        ByVal sIOMode _
    )
        Dim lIOMode
        Select Case LCase( sIOMode )
            Case "r" :  lIOMode = 1
            Case "w" :  lIOMode = 2
            Case "+w" : lIOMode = 8
            Case Else : Exit Function
        End Select
        
        Set goLogFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sTrgtPath, lIOMode, True)
        gbLogFileEnable = True
    End Function
    
    ' ==================================================================
    ' = �T�v    ���O����������
    ' = ����    sOutputMsg  String   [in]   �o�̓��b�Z�[�W
    ' = �ߒl    �Ȃ�
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Function Puts( _
        ByVal sOutputMsg _
    )
        If gbLogFileEnable = True Then
            goLogFile.WriteLine sOutputMsg
        Else
            WScript.Echo sOutputMsg
        End If
    End Function
    
    ' ==================================================================
    ' = �T�v    ���O�o�͂��~����
    ' = ����    �Ȃ�
    ' = �ߒl    �Ȃ�
    ' = �o��    �{�֐����Ăяo���ƁA�o�̓��[�h�������I�Ɂu�W���o�́v��
    ' =         �؂�ւ��
    ' ==================================================================
    Public Function Close()
        If gbLogFileEnable = True Then
            goLogFile.Close
            gbLogFileEnable = False
        Else
            'Do Nothing
        End If
    End Function
End Class
    If WScript.ScriptName = "Log.vbs" Then
        Call Test_LogMng
    End If
    Private Sub Test_LogMng
        Dim oLog1
        Set oLog1 = New LogMng
        Dim oLog2
        Set oLog2 = New LogMng
        Call oLog1.Open( Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" ) & "\LogTest1.log", "+w" )
        Call oLog2.Open( Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" ) & "\LogTest2.log", "+w" )
        
        oLog1.Puts "desu"
        oLog1.Puts "yorosiku"
        oLog2.Puts "desu"
        oLog2.Puts "yorosiku"
        oLog2.Puts "you"
        Call oLog2.Close
        oLog2.Puts "desu"
        oLog2.Puts "yorosiku"
        
        Call oLog1.Close
        Call oLog2.Close
        Set oLog1 = Nothing
        Set oLog2 = Nothing
    End Sub
