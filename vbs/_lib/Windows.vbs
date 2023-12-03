Option Explicit

' ==================================================================
' = �T�v    �Ǘ��Ҍ����Ŏ��s����i���X�N���v�g�����s�j
' = ����    �Ȃ�
' = �ߒl                Boolean     [out]   ���s����
' = �o��    �E�����X�N���v�g���Ǘ��Ҍ����ŊJ�������Ď��s����B
' =           �g������͗ǂ����A���p�\�Ȗ��߂�����Ȃ�
' =           �\�t�g�E�F�A�i��:X-Finder)�ł͎g�p�ł��Ȃ��B
' =         �E�{�֐����Ăяo���X�N���v�g���Ŏ󂯎���������́A
' =           �{�֐����o�R����Ɨ��p�ł��Ȃ��Ȃ邽�ߗv���ӁB
' =           ����������肵�����ꍇ�́A�e�L�X�g�t�@�C���ǂݏ����𗘗p���邱�ƁB
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExecRunas()
    Dim oArgs
    Dim bIsRunas
    Dim sArgs
    
    bIsRunas = False
    sArgs = ""
    Set oArgs = WScript.Arguments
    
    ' �t���O�̎擾
    If oArgs.Count > 0 Then
        If UCase(oArgs.item(0)) = "/RUNAS" Then
            bIsRunas = True
        End If
        sArgs = sArgs & " " & oArgs.item(0)
    End If
    
    Dim bIsExecutableOs
    bIsExecutableOs = false
    Dim oOsInfos
    Set oOsInfos = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
    Dim oOs
    For Each oOs in oOsInfos
        If Left(oOs.Version, 3) >= 6.0 Then
            bIsExecutableOs = True
        End If
    Next
    
    Dim oWshShell
    Set oWshShell = CreateObject("Shell.Application")
    ExecRunas = False
    If bIsRunas = False Then
        If bIsExecutableOs = True Then
            oWshShell.ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " /RUNAS " & sArgs, "", _
            "runas", _
            1
            ExecRunas = True
            Wscript.Quit
        End If
    End If
End Function

' ==================================================================
' = �T�v    �Ǘ��Ҍ����Ŏ��s����i�ʃX�N���v�g���s�j
' = ����    sScriptPath     String      [in]    �X�N���v�g�t�@�C���p�X
' = �ߒl    �Ȃ�
' = �o��    �E�ʂ̃X�N���v�g���Ǘ��Ҍ����ŊJ���Ď��s����B
' =           ��ɗ��p�\�Ȗ��߂�����Ȃ��\�t�g�E�F�A�i��:X-Finder�j
' =           �ɂė��p���邱�Ƃ�z�肵�Ă���B
' =         �E�{�֐����Ăяo���X�N���v�g���Ŏ󂯎���������́A
' =           �{�֐����o�R����Ɨ��p�ł��Ȃ��Ȃ邽�ߗv���ӁB
' =           ����������肵�����ꍇ�́A�e�L�X�g�t�@�C���ǂݏ����𗘗p���邱�ƁB
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExecRunas2( _
    ByVal sScriptPath _
)
    Dim objShell
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute "wscript.exe", sScriptPath & " runas", "", "runas", 1
End Function

' ==================================================================
' = �T�v    Dos �R�}���h���s
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExecDosCmd( _
    ByVal sCommand _
)
    Dim oExeResult
    Dim sStrOut
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
    Do While Not (oExeResult.StdOut.AtEndOfStream)
        sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
'   Call Test_ExecDosCmd()
    Private Sub Test_ExecDosCmd()
        Msgbox ExecDosCmd( "copy ""C:\Users\draem_000\Desktop\test.txt"" ""C:\Users\draem_000\Desktop\test2.txt""" )
        'Msgbox ExecDosCmd( "C:\codes\vbs\_lib\test.bat" )
    End Sub

' ==================================================================
' = �T�v    �v���Z�X�N���m�F
' = ����    sProcessName    String      [in]    �v���Z�X��
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExistProcess( _
    ByVal sProcessName _
)
    Dim objService
    Dim objQfeSet
    Set objService = CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set objQfeSet = objService.ExecQuery("Select * From Win32_Process Where Caption Like '" & sProcessName & "%'")
    ExistProcess = objQfeSet.Count > 0
End Function
'   Call Test_ExistProcess()
    Private Sub Test_ExistProcess()
        MsgBox ExistProcess("wsl.exe")
    End Sub

' ==================================================================
' = �T�v    WSL2 Running �҂�
' = ����    sDistName   String  [in]    WSL2 �f�B�X�g���r���[�V������
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Private Function WaitForWslRunning( _
    ByVal sDistName _
)
    Const sLOG_FILE_NAME = "wsl_status.log"
    
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    
    Dim sLogFilePath
    sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sLOG_FILE_NAME
    
    Dim sStatus
    Do
        On Error Resume Next
        WScript.sleep(500)
        
        ' wsl�X�e�[�^�X�R�}���h ���_�C���N�g
        objWshShell.Run "%comspec% /c wsl -l -v > """ & sLogFilePath & """", 0 , True
        
        '���_�C���N�g���O�ǂݍ���
        Dim adoStrm
        Set adoStrm = CreateObject("ADODB.Stream")
        adoStrm.Type = 2
        adoStrm.Charset = "UTF-16"
        adoStrm.LineSeparator = -1
        adoStrm.Open
        adoStrm.LoadFromFile sLogFilePath
        Dim sLine
        Dim sStatusLine
        Do Until adoStrm.EOS
            sLine = adoStrm.ReadText(-2)
            If InStr(sLine, "* " & sDistName) Then
                sStatusLine = sLine
            End If
        Loop
        
        '�X�e�[�^�X�擾
        Dim sTargetStr
        sTargetStr = sLine
        Dim sSearchPattern
        sSearchPattern = "^\*\s+(" & sDistName & ")\s+(\w+)"
        oRegExp.IgnoreCase = True
        oRegExp.Global = True
        oRegExp.Pattern = sSearchPattern
        Dim oMatchResult
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        sStatus = oMatchResult(0).SubMatches(1)
        'MsgBox sStatus
        On Error Goto 0
    Loop While sStatus <> "Running"
    objFSO.DeleteFile sLogFilePath, True
End Function

