Option Explicit

'<<概要>>
'  指定した時刻までスリープを抑止し、その時刻にWindowsをシャットダウンする（退勤予約）
'
'<<使用方法>>
'  1) ScheduleClockOut.vbs を実行する
'  2) シャットダウン時刻を HH:MM 形式（例: 19:01）で入力する
'  3) Shutdown.vbs の確認ダイアログで OK を押すと予約が開始される
'  ※ 引数で時刻(HH:MM)を渡した場合は入力を省略できる
'
'<<仕様>>
'  ・PeriodicKeyTransmission.vbs -f を起動してスリープを抑止する
'  ・Shutdown.vbs <時刻> を起動してシャットダウンを予約する
'  ・連携スクリプトは本スクリプトと同一フォルダに配置されていること
'  ・PeriodicKeyTransmission.vbs は多重起動時に先行プロセスを停止する仕様のため、
'    既にキー送信中の場合は起動によってキー送信が停止する
'  ・Shutdown.vbs の確認ダイアログをキャンセルした場合はキー送信のみが残るため、
'    停止するには PeriodicKeyTransmission.vbs を再実行する

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "退勤予約"
Const sDEFAULT_TIME = "19:05"
Const sKEY_SCRIPT_NAME = "PeriodicKeyTransmission.vbs"
Const sSHUTDOWN_SCRIPT_NAME = "Shutdown.vbs"

'===============================================================================
'= 本処理部
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= メイン関数
'===============================================================================
Public Sub Main()
    '*************************************************
    '* 時刻入力（HH:MM）
    '*************************************************
    Dim sInputTime
    If cArgs.Count >= 1 Then
        sInputTime = cArgs(0)
    Else
        sInputTime = InputBox( "シャットダウン時刻を HH:MM 形式（例: 19:01）で入力してください", sSCRIPT_NAME, sDEFAULT_TIME )
    End If

    sInputTime = Trim(sInputTime)
    If sInputTime = "" Then
        'キャンセル
        Exit Sub
    End If

    If IsValidTimeFormat(sInputTime) = False Then
        MsgBox "時刻の形式が正しくありません（HH:MM形式で入力してください）。処理を中断します。", vbExclamation, sSCRIPT_NAME
        Exit Sub
    End If

    '*************************************************
    '* 連携スクリプトのパス解決
    '*************************************************
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim sScriptDir
    sScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

    Dim sKeyScriptPath, sShutdownScriptPath
    sKeyScriptPath = objFSO.BuildPath(sScriptDir, sKEY_SCRIPT_NAME)
    sShutdownScriptPath = objFSO.BuildPath(sScriptDir, sSHUTDOWN_SCRIPT_NAME)

    If objFSO.FileExists(sKeyScriptPath) = False Then
        MsgBox sKeyScriptPath & " が見つかりません。処理を中断します。", vbExclamation, sSCRIPT_NAME
        Exit Sub
    End If
    If objFSO.FileExists(sShutdownScriptPath) = False Then
        MsgBox sShutdownScriptPath & " が見つかりません。処理を中断します。", vbExclamation, sSCRIPT_NAME
        Exit Sub
    End If

    '*************************************************
    '* 連携スクリプト起動
    '*************************************************
    'いずれも常駐するため、戻りを待たずに起動する
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    objWshShell.Run "wscript.exe """ & sKeyScriptPath & """ -f", 0, False
    objWshShell.Run "wscript.exe """ & sShutdownScriptPath & """ " & sInputTime, 0, False
End Sub

'===============================================================================
'= 内部関数
'===============================================================================
'-------------------------------------------------------------------------------
'- 時刻形式（HH:MM）の妥当性判定
'-------------------------------------------------------------------------------
Private Function IsValidTimeFormat( ByVal sTime ) '{{{
    IsValidTimeFormat = False

    Dim aParts
    aParts = Split(sTime, ":")
    If UBound(aParts) <> 1 Then
        Exit Function
    End If

    If IsNumeric(aParts(0)) = False Or IsNumeric(aParts(1)) = False Then
        Exit Function
    End If

    Dim lHour, lMinute
    lHour = CInt(aParts(0))
    lMinute = CInt(aParts(1))
    If lHour < 0 Or lHour > 23 Then
        Exit Function
    End If
    If lMinute < 0 Or lMinute > 59 Then
        Exit Function
    End If

    IsValidTimeFormat = True
End Function '}}}

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1

    MsgBox "=== test start ==="

    Select Case lTestCase
        Case 1
            cArgs.Add "19:01"
            Call Main()
            MsgBox "1 実行後"
        Case Else
            Call Main()
    End Select

    MsgBox "=== test finished ==="
End Sub '}}}
