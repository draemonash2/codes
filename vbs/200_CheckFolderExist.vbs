Option Explicit
Dim sDirPath
sDirPath = "C:\Users\draem_000\AppData\Local\Temp\msupdate71"

Dim sNow
Dim sCmpBaseTime
sNow = Now()
sCmpBaseTime = InputBox( _
                    "�I����������" & vbNewLine & _
                    "" & vbNewLine & _
                    "  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
                    "" & vbNewLine & _
                    "�� ���t�݂̂��w�肵�����ꍇ�A�uYYYY/MM/DD 0:0:0�v�Ƃ��Ă��������B" _
                    , "����" _
                    , sNow _
                )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Do While 1
    If DateDiff( "s", sCmpBaseTime, Now() ) < 0 Then
        If objFSO.FolderExists( sDirPath ) Then
            MsgBox "�t�H���_���쐬����܂����I" & vbNewLine & sDirPath
            MsgBox "�v���O�������I�����܂��I"
            WScript.Quit
        Else
            'Do Nothing
        End If
    Else
        MsgBox "�I�������ɂȂ�܂����I"
        MsgBox "�v���O�������I�����܂��I"
        WScript.Quit
    End If
    WScript.Sleep 5000
Loop
