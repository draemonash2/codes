Option Explicit
Dim sDirPath
sDirPath = "C:\Users\draem_000\AppData\Local\Temp\msupdate71"

Dim sNow
Dim sCmpBaseTime
sNow = Now()
sCmpBaseTime = InputBox( _
                    "終了日時入力" & vbNewLine & _
                    "" & vbNewLine & _
                    "  [入力規則] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
                    "" & vbNewLine & _
                    "※ 日付のみを指定したい場合、「YYYY/MM/DD 0:0:0」としてください。" _
                    , "入力" _
                    , sNow _
                )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Do While 1
    If DateDiff( "s", sCmpBaseTime, Now() ) < 0 Then
        If objFSO.FolderExists( sDirPath ) Then
            MsgBox "フォルダが作成されました！" & vbNewLine & sDirPath
            MsgBox "プログラムを終了します！"
            WScript.Quit
        Else
            'Do Nothing
        End If
    Else
        MsgBox "終了時刻になりました！"
        MsgBox "プログラムを終了します！"
        WScript.Quit
    End If
    WScript.Sleep 5000
Loop
