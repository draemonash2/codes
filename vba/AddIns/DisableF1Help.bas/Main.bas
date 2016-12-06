Attribute VB_Name = "Main"
Option Explicit

Const APPNAME = "F1ÇÃÉwÉãÉvÇñ≥å¯Ç…"

Const MSG_OUTPUT_SECOND As Long = 10

Public Sub InstallAddin()
    Call ModifyRegistry(True)
End Sub

Public Sub UninstallAddin()
    Call ModifyRegistry(False)
End Sub

Private Sub ModifyRegistry( _
    ByVal bAdd As Boolean _
)
    Dim tRegStruct As T_REG_STRUCT
    ReDim Preserve tRegStruct.atRegKeys(0)
    ReDim Preserve tRegStruct.atRegKeys(0).atRegValues(0)
    
    On Error Resume Next
    With tRegStruct
        With .atRegKeys(0)
            .sKey = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\DisabledShortcutKeysCheckBoxes"
            .eOperation = REG_ADDMOD
            If bAdd = True Then
                With .atRegValues(0)
                    .sName = "F1key"
                    .sData = "112,0"
                    .eType = REG_SZ
                    .eOperation = REG_ADDMOD
                End With
            Else
                With .atRegValues(0)
                    .sName = "F1key"
                    .eOperation = REG_DELETE
                End With
            End If
        End With
    End With
    Call EnableRegWrite
    Call SetRegistry(APPNAME, tRegStruct)
    On Error GoTo 0
End Sub

Private Sub auto_open()
    On Error Resume Next
    Call SetStatusBar
    On Error GoTo 0
End Sub

Private Sub SetStatusBar()
    Dim sMsg As String
    Dim tvWakeupTime As Variant
    Dim tvMsgOutputSec As Variant
    Dim tvNow As Variant
    
    sMsg = "F1ÇÃÉwÉãÉvÇÕñ≥å¯Ç…ê›íËÇ≥ÇÍÇƒÇ¢Ç‹Ç∑ÅB"
    tvMsgOutputSec = TimeValue(CStr(CDate(MSG_OUTPUT_SECOND / 86400#)))
    tvNow = TimeValue(Format(Now, "hh:mm:ss"))
    tvWakeupTime = tvNow + tvMsgOutputSec
    
    Application.StatusBar = sMsg
    Application.OnTime tvWakeupTime, "'ClearStatusBar'"
End Sub

Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

