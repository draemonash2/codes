Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.02

' ==================================================================
' = 概要    コマンドを実行
' = 引数    sCommand    String   [in]   コマンド
' = 引数    bGetStdout  Boolean  [in]   標準出力取得有無(省略可)
' = 戻値                String          標準出力
' = 覚書    ・大量の処理を行うbatを実行する場合、bGetStdoutをFalseにすること。
' =           コマンドの実行結果が必要な場合は、コマンドにリダイレクトを含めること。
' =             例）Call ExecDosCmd("xxx.bat > xxx.log", False)
' =           【理由】
' =           Execは標準出力にためるバッファの最大は4096バイトであり、
' =           それ以上のデータを読み込むとAtEndOfStream時に固まるため。
' =           https://community.cybozu.dev/t/topic/181/2
' = 依存    なし
' = 所属    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmd( _
    ByVal sCommand As String, _
    Optional bGetStdOut As Boolean = True _
) As String
    If sCommand = "" Then
        ExecDosCmd = ""
    Else
        Dim sStdOutAll As String
        sStdOutAll = ""
        If bGetStdOut = True Then
            Dim oExeResult As Object
            Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
            Do While Not oExeResult.StdOut.AtEndOfStream
                Dim sStdOut As String
                sStdOut = oExeResult.StdOut.ReadLine
                Debug.Print sStdOut
                sStdOutAll = sStdOutAll & vbNewLine & sStdOut
            Loop
            Set oExeResult = Nothing
        Else
            Call CreateObject("WScript.Shell").Run("%ComSpec% /c """ & sCommand & """", WaitOnReturn:=True)
        End If
        ExecDosCmd = sStdOutAll
    End If
End Function
    Private Sub Test_ExecDosCmd()
        Dim sBuf As String
        sBuf = sBuf & vbNewLine & ExecDosCmd("copy C:\Users\draem_000\Desktop\test.txt C:\Users\draem_000\Desktop\test2.txt")
        MsgBox sBuf
    End Sub

' ==================================================================
' = 概要    コミットダイアログを表示
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 依存    Mng_SysCmd.bas/ExecDosCmd()
' = 所属    Mng_SysCmd.bas
' ==================================================================
Private Function ShowCommitDialog()
    Dim sCmdRslt As String
    Dim sCmd As String
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "★" & """ "
    Else
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "★" & """ "
    End If
    sCmdRslt = ExecDosCmd(sCmd)
End Function

