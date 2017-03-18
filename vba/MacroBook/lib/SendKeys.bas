Attribute VB_Name = "SendKeys"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SendKeysBetweenWait( _
    ByVal sSendKeys As String, _
    ByVal lWaitTime As Long _
)
    Dim lIdx As Long
    Dim bIsConc As Boolean
    Dim bIsExec As Boolean
    Dim sConcStr As String
    Dim sChar As String
    sConcStr = ""
    bIsConc = False
    For lIdx = 1 To Len(sSendKeys)
        sChar = Mid$(sSendKeys, lIdx, 1)
        'If sChar = "^" Or sChar = "+" Or sChar = "%" Then
        If sChar = "^" Or sChar = "+" Then
            sConcStr = sConcStr & sChar
            bIsExec = False
        ElseIf sChar = "{" Then
            sConcStr = sConcStr & sChar
            bIsConc = True
            bIsExec = False
        ElseIf sChar = "}" Then
            sConcStr = sConcStr & sChar
            bIsConc = False
            bIsExec = True
        Else
            sConcStr = sConcStr & sChar
            If bIsConc = True Then
                bIsExec = False
            Else
                bIsExec = True
            End If
        End If
        If bIsExec = True Then
'            Application.SendKeys sConcStr, True
'            Sleep (lWaitTime)
            Debug.Print sConcStr
            sConcStr = ""
        Else
            'Do Nothing
        End If
    Next lIdx
End Sub
    Private Sub Test_SendKeysBetweenWait()
        Call SendKeysBetweenWait("%^+{hd}+d", 10)
        Call SendKeysBetweenWait("%d", 10)
    End Sub



