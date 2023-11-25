Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.03

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'for ExecDosCmdRunas()
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long 'for ExecDosCmdRunas()

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
' = 概要    コマンドを実行（管理者権限）
' = 引数    asCommands()    String   [in] 実行コマンド
' = 引数    bDelFiles       Boolean  [in] Bat/Logファイル削除(省略可)
' = 戻値                    String        標準出力＆標準エラー出力
' = 覚書    ・Desktopフォルダパスに空白が含まれる場合は、動作しない。
' = 依存    なし
' = 依存    Mng_FileSys.bas/OutputTxtFile()
' = 所属    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmdRunas( _
    ByRef asCommands() As String, _
    Optional bDelFiles As Boolean = True _
) As String
    Const sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME As String = "CmdExeBatRunas"
    If Sgn(asCommands) = 0 Then
        ExecDosCmdRunas = ""
    Else
        If UBound(asCommands) < 0 Then
            ExecDosCmdRunas = ""
        Else
            Dim objWshShell
            Set objWshShell = CreateObject("WScript.Shell")
            Dim objFSO
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            
            Dim sBatFilePath As String
            Dim sLogFilePath As String
            sBatFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".bat"
            sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".log"
            
            '「@echo off」挿入
            ReDim Preserve asCommands(UBound(asCommands) + 1)
            Dim lIdx As Long
            For lIdx = UBound(asCommands) To (LBound(asCommands) + 1) Step -1
                asCommands(lIdx) = asCommands(lIdx - 1)
            Next lIdx
            asCommands(0) = "@echo off"
            
            'BATファイル作成
            Call OutputTxtFile(sBatFilePath, asCommands)
            Do While Not objFSO.FileExists(sBatFilePath)
                Sleep 100
            Loop
            
            'BATファイル実行
            ShellExecute 0, "runas", sBatFilePath, " > " & sLogFilePath & " 2>&1", vbNullString, 1
            
            'LOGファイル出力待ち
            Do While Not objFSO.FileExists(sLogFilePath)
                Sleep 100
            Loop
            
            'LOGファイル読込み
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(sLogFilePath, 1, True)
            Dim sStdOutAll As String
            sStdOutAll = ""
            Dim sLine As String
            Do Until objTxtFile.AtEndOfStream
                sLine = objTxtFile.ReadLine
                'MsgBox sLine
                If sStdOutAll = "" Then
                    sStdOutAll = sLine
                Else
                    sStdOutAll = sStdOutAll & vbNewLine & sLine
                End If
            Loop
            'MsgBox sStdOutAll
            objTxtFile.Close
            
            'BATファイル/LOGファイル削除
            If bDelFiles = True Then
                Kill sBatFilePath
                Kill sLogFilePath
            End If
            
            ExecDosCmdRunas = sStdOutAll
        End If
    End If
End Function
    Private Sub Test_ExecDosCmdRunas()
        Dim asCommands() As String
        
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(0)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source2.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source2.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands, False)
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

