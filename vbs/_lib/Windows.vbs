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
