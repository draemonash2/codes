Option Explicit

' ==================================================================
' = �T�v    �f�o�b�O�p�̏o�͏���(Collection�p)
' = ����    cCollection     Collections [in]    �o�͑Ώ�
' = �ߒl    �Ȃ�
' = �o��    �E�o�͌�͏������I������
' = �ˑ�    �Ȃ�
' = ����    Debug.vbs
' ==================================================================
Private Function DebugPrintClct( _
    ByRef cCollection _
)
    Dim sDebugMsg
    Dim vValue
    For Each vValue In cCollection
        sDebugMsg = sDebugMsg & vbNewLine & vValue
    Next
    MsgBox sDebugMsg
    WScript.Quit
End Function
