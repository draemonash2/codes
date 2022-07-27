Attribute VB_Name = "Mng_Svn"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private gbIsForceExec As Boolean '強制コミットモード（エラー有無に関わらず、コミット処理を実施させる）

Public Enum E_SVN_MOD_STATUS
    MOD_STAT_OUTOFVERCTRL = 0
    MOD_STAT_NOTCHANGE
    MOD_STAT_ADDED
    MOD_STAT_MODIFIED
    MOD_STAT_OTHER
End Enum

Public Type T_SVN_MOD_STAT_INFO
    sPath As String
    eSvnModStat As E_SVN_MOD_STATUS
End Type

Public Function SvnInit()
    gbIsForceExec = False
End Function

Public Function ExecCommit()
    Dim bIsCommitExec As Boolean

    '=== コミット前チェック処理
    bIsCommitExec = PreChkBeforeCommit()
    
    '=== コミット処理 ===
    If bIsCommitExec = True Then
        Call AddFiles2Svn
        Call ShowCommitDialog
    Else
        Call OutpErrorMsg(ERROR_PROC_STOP) 'エラー出力後、処理停止
    End If
End Function

Private Function PreChkBeforeCommit() As Boolean
    Dim bIsCommitExec As Boolean
    Dim lProcSel As Long
    Dim bChkProcFin As Boolean
    
    bIsCommitExec = True
    bChkProcFin = False
    
    '=== エラー発生確認 ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If gbIsForceExec = True Then '強制コミットモードが ON の時はエラーを無視する！
            bChkProcFin = True
            bIsCommitExec = True
        Else
            If ErrorExist() = False Then
                'Do Nothing
            Else
                Call StoreErrorMsg( _
                    "エラーが発生しています！" & vbNewLine & _
                    "項目書もしくはログファイルを確認してください！" _
                )
                bChkProcFin = True
                bIsCommitExec = False
            End If
        End If
    End If
    
    '=== ワーニング発生確認 ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If gbIsForceExec = True Then '強制コミットモードが ON の時はワーニングを無視する！
            bChkProcFin = True
            bIsCommitExec = True
        Else
            If WarningExist() = False Then
                'Do Nothing
            Else
                lProcSel = MsgBox( _
                                    "ワーニングが発生していますが、処理を続けますか？" & vbNewLine & _
                                    "" & vbNewLine & _
                                    "ワーニングの内容を確認したい場合は「キャンセル」を、" & vbNewLine & _
                                    "既知の問題であれば「OK」を押してください。", _
                                    vbOKCancel _
                           )
                If lProcSel = vbOK Then
                    'Do Nothing
                Else
                    Call StoreErrorMsg("処理がキャンセルされました！")
                    bChkProcFin = True
                    bIsCommitExec = False
                End If
            End If
        End If
    End If
    
    '=== CUI 版 SVN インストール確認 ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        If ChkCuiSvnInstall() = True Then
            'Do Nothing
        Else
            Call StoreErrorMsg( _
                "CUI 版 Subversion がインストールされておりません！" & vbNewLine & _
                "インストール後に再度実行してください！" _
            )
            bChkProcFin = True
            bIsCommitExec = False
        End If
    End If
    
    '=== 最終確認 ===
    If bChkProcFin = True Then
        'Do Nothing
    Else
        MsgBox "ログフォルダ構成・ログファイル名に問題がなかったため、" & vbNewLine & _
               "ログファイルを SVN ADD してコミットダイアログを表示します。"
        lProcSel = MsgBox( _
                            "試験ログファイルを SVN ADD します。" & vbNewLine & _
                            "問題がなければ「OK」を押してください。" & vbNewLine & _
                            "" & vbNewLine & _
                            "（ADD 済みの場合も「OK」としてください）", _
                            vbOKCancel _
                   )
        If lProcSel = vbOK Then
            'Do Nothing
        Else
            Call StoreErrorMsg("処理がキャンセルされました！")
            bChkProcFin = True
            bIsCommitExec = False
        End If
    End If
    
    PreChkBeforeCommit = bIsCommitExec
End Function

Private Function ChkCuiSvnInstall() As Boolean
    Dim sCmdRslt As String
    sCmdRslt = ExecDosCmd("svn --version")
    'Debug.Print sCmdRslt
    If InStr(sCmdRslt, "svn, version") > 0 And _
       InStr(sCmdRslt, "Copyright") > 0 Then
        ChkCuiSvnInstall = True
    ElseIf sCmdRslt = "" Then
        ChkCuiSvnInstall = False
    Else
        Stop
    End If
End Function

Private Function AddFiles2Svn()
    Dim sCmdRslt As String
    Dim sCmd As String
    
    '*** ログフォルダ配下 ***
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "svn add --parents " & gtInputInfo.sTestLogDirPath
    Else
        sCmd = "svn add --parents " & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName
    End If
    sCmdRslt = ExecDosCmd(sCmd)
    '既に追加済みの場合、sCmdRslt は "" となってしまうため、
    'エラーチェックしない。
'    If sCmdRslt = "" Then
'        Call StoreErrorMsg( _
'            "svn add コマンドが失敗しました！" _
'        )
'        Call OutpErrorMsg(ERROR_PROC_STOP)
'    Else
'        'Do Nothing
'    End If
    
    '*** 試験項目書 ***
    sCmd = "svn add --parents " & gtInputInfo.sTestDocFilePath
    sCmdRslt = ExecDosCmd(sCmd)
    '既に追加済みの場合、sCmdRslt は "" となってしまうため、
    'エラーチェックしない。
'    If sCmdRslt = "" Then
'        Call StoreErrorMsg( _
'            "svn add コマンドが失敗しました！" _
'        )
'        Call OutpErrorMsg(ERROR_PROC_STOP)
'    Else
'        'Do Nothing
'    End If
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

Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    'Debug.Print sStrOut
    'Debug.Print ""
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function

Public Function GetSvnModStatList( _
    ByVal sFilePath As String _
) As T_SVN_MOD_STAT_INFO()
    Dim sCmd As String
    Dim sCmdRslt As String
    Dim asCmdRslt() As String
    Dim lCmdRsltCnt As Long
    Dim sExtFilePath As String
    Dim sModStatus As String
    Dim eModStatus As E_SVN_MOD_STATUS
    Dim sCmdRsltLine As String
    Dim atSvnModStat() As T_SVN_MOD_STAT_INFO
    Dim lSvnModStatIdx As Long
    Const E_PATH_STRT_POS = 42
    
    sCmd = "svn status -v " & sFilePath
    sCmdRslt = ExecDosCmd(sCmd)
    asCmdRslt = Split(sCmdRslt, vbNewLine)
    For lCmdRsltCnt = 0 To UBound(asCmdRslt)
        sCmdRsltLine = asCmdRslt(lCmdRsltCnt)
        If sCmdRsltLine = "" Then
            'Do Nothing
        Else
            If Sgn(atSvnModStat) = 0 Then
                lSvnModStatIdx = 0
            Else
                lSvnModStatIdx = lSvnModStatIdx + 1
            End If
            ReDim Preserve atSvnModStat(lSvnModStatIdx)
            sModStatus = Mid$(sCmdRsltLine, 1, 1)
            sExtFilePath = Mid$(sCmdRsltLine, E_PATH_STRT_POS, Len(sCmdRsltLine) - E_PATH_STRT_POS + 1)
            Select Case sModStatus
                Case " ": eModStatus = MOD_STAT_NOTCHANGE
                Case "?": eModStatus = MOD_STAT_OUTOFVERCTRL
                Case "A": eModStatus = MOD_STAT_ADDED
                Case "M": eModStatus = MOD_STAT_MODIFIED
                Case Else: eModStatus = MOD_STAT_OTHER
            End Select
            atSvnModStat(lSvnModStatIdx).sPath = sExtFilePath
            atSvnModStat(lSvnModStatIdx).eSvnModStat = eModStatus
            'Debug.Print oSvnModStatList.Item(sExtFilePath) & ":" & sExtFilePath
        End If
    Next lCmdRsltCnt
    GetSvnModStatList = atSvnModStat
End Function

'★テスト用★
Sub test3()
    Dim sFilePath As String
    Dim lLoopCnt As Long
    Dim oStopWatch As New StopWatch
    Dim atSvnModStat() As T_SVN_MOD_STAT_INFO
    
    sFilePath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\Busmaster"
    
    Call SvnInit
    Call oStopWatch.StartT
    atSvnModStat = GetSvnModStatList(sFilePath)
    Debug.Print oStopWatch.StopT
End Sub

Sub test()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\commit_test\ST"
    gtInputInfo.sTestDocFilePath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\test\TF2次プロト2ndシステム試験項目書-17650_○.xlsx"
    Call ShowCommitDialog
End Sub

Sub test2()
    Call InputInfoInit
    gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\commit_test\ST"
    'gtInputInfo.sTestLogDirPath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\commit_test"
    gtInputInfo.sTestDocFilePath = "C:\Coffer\svn_PTM\trunk\03_管理\30_ツール\test\TF2次プロト2ndシステム試験項目書-17650_○.xlsx"
    Call AddFiles2Svn
End Sub

