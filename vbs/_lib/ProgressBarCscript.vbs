Option Explicit
'�Q�lURL:http://d.hatena.ne.jp/maeyan/touch/20140416/1397602663

' progrress bar cscript class v1.00

' = �ˑ�    �Ȃ�
' = ����    ProgressBarCscript.vbs
Class ProgressBar
    Private sStatus
    Private objFSO
    Private objWshShell
    
    Private Sub Class_Initialize
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        Dim sExeFileName
        sExeFileName = LCase(objFSO.GetFileName(WScript.FullName))
        if sExeFileName = "cscript.exe" then
            'Do Nothing
        else
            objWshShell.Run "cscript //nologo """ & Wscript.ScriptFullName & """", 1, False
            Wscript.Quit
        end if
    End Sub
    
    Private Sub Class_Terminate
        Set objFSO = Nothing
        Set objWshShell = Nothing
    End Sub
    
    ' ==================================================================
    ' = �T�v    ���b�Z�[�W���X�V����
    ' = ����    sProgMsg      String   [in] ���b�Z�[�W
    ' = �ߒl    �Ȃ�
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Property Let Message( _
        ByVal sMessage _
    )
        if sStatus = "Update" then
            Wscript.StdOut.Write vbCrLf
        end if
        Wscript.StdOut.Write sMessage & vbCrLf
        sStatus = "Message"
    End Property
    
    ' ==================================================================
    ' = �T�v    �i�����X�V����
    ' = ����    lBunsi      Long   [in] �i��
    ' = ����    lBunbo      Long   [in] �i���ő�l
    ' = �ߒl    �Ȃ�
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Sub Update( _
        ByVal lBunsi, _
        ByVal lBunbo _
    )
        '�p�[�Z���e�[�W�v�Z
        Dim iPercentage
        Dim sPercentage
        iPercentage = Cint((lBunsi / lBunbo) * 100)
        sPercentage = iPercentage & "%"
        sPercentage = String(4 - Len(sPercentage), " ") & sPercentage
        
        '�i���o�[
        Dim sProgressBar
        sProgressBar = String(Cint(iPercentage/5), "=") & ">" & String(20 - Cint(iPercentage/5), " ")
        
        '�`��
        Wscript.StdOut.Write sPercentage & " |" & sProgressBar & "| " & lBunsi & "/" & lBunbo & vbCr
        sStatus = "Update"
    End Sub
    
    ' ==================================================================
    ' = �T�v    �v���O���X�o�[���I������
    ' = ����    �Ȃ�
    ' = �ߒl    �Ȃ�
    ' = �o��    cscript�͏I���ł��Ȃ�
    ' ==================================================================
'   Public Function Quit()
'       gobjExplorer.Document.Body.Style.Cursor = "default"
'       gobjExplorer.Quit
'   End Function
    
End Class
    If WScript.ScriptName = "ProgressBarCscript.vbs" Then
        Call Test_ProgressBar
    End If
    Private Sub Test_ProgressBar
        Dim lProcIdx
        Dim lProcNum
        Dim objPrgrsBar
        Set objPrgrsBar = New ProgressBar
        
        '#�����P
        objPrgrsBar.Message = "�������� ���s!"
        lProcNum = 255
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 1
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        '#�����Q
        objPrgrsBar.Message = "�Z������ ���s!"
        lProcNum= 10
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 45
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        objPrgrsBar.Message = "Complete!!"
        msgbox "�I�����܂���"
    End Sub
