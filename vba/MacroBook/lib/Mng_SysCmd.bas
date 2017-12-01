Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.0

Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
 
Private Function ShowCommitDialog()
    Dim sCmdRslt As String
    Dim sCmd As String
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "Åö" & """ "
    Else
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "Åö" & """ "
    End If
    sCmdRslt = ExecDosCmd(sCmd)
End Function

