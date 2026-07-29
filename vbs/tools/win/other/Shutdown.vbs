Option Explicit

'<<概要>>
'  指定した時刻にWindowsをシャットダウンする
'
'<<使用方法>>
'  1) Shutdown.vbs を実行する
'  2) シャットダウン時刻を HH:MM 形式（例: 19:01）で入力する
'  3) 確認ダイアログで OK を押すと予約が開始される
'  ※ 引数で時刻(HH:MM)を渡した場合は入力を省略できる
'
'<<仕様>>
'  ・入力時刻が現在時刻より前の場合は翌日の同時刻を対象とする
'  ・予約開始時に実行中ファイルをデスクトップに出力し、完了時に削除する

'===============================================================================
'= インクルード
'===============================================================================
'Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "シャットダウン予約"
Const bOUTPUT_EXECFILE = False

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
        sInputTime = InputBox( "シャットダウン時刻を HH:MM 形式（例: 19:01）で入力してください", sSCRIPT_NAME )
        If IsEmpty(sInputTime) = True Then
            'キャンセル
            Exit Sub
        End If
    End If

    sInputTime = Trim(sInputTime)
    If IsValidTimeFormat(sInputTime) = False Then
        MsgBox "時刻の形式が正しくありません（HH:MM形式で入力してください）。処理を中断します。", vbExclamation, sSCRIPT_NAME
        Exit Sub
    End If

    '*************************************************
    '* 対象日時算出
    '*************************************************
    Dim dTargetTime
    dTargetTime = CalcTargetDateTime(sInputTime)

    Dim vAnswer
    vAnswer = MsgBox( FormatDateTime(dTargetTime) & " にシャットダウンします。よろしいですか？", vbOkCancel + vbQuestion, sSCRIPT_NAME )
    If vAnswer <> vbOk Then
        Exit Sub
    End If

    '*************************************************
    '* シェルオブジェクト生成
    '*************************************************
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")

    '*************************************************
    '* 実行中ファイル出力
    '*************************************************
    Dim objFSO
    Dim sRunFilePath
    If bOUTPUT_EXECFILE Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        'ファイル名に使用できない ":" は代替データストリーム扱いとなるため置換する
        sRunFilePath = objWshShell.SpecialFolders("Desktop") & "\running_shutdown_[" & Replace(sInputTime, ":", "-") & "].txt"
        On Error Resume Next
        Dim objTxtFile
        Set objTxtFile = objFSO.OpenTextFile(sRunFilePath, 2, True)
        objTxtFile.WriteLine sInputTime & " にシャットダウン予約中..."
        objTxtFile.Close
        On Error Goto 0
    End If

    '*************************************************
    '* 待機
    '*************************************************
    Dim lWaitSec
    lWaitSec = DateDiff("s", Now(), dTargetTime)
    If lWaitSec > 0 Then
        WScript.Sleep( lWaitSec * 1000 )
    End If

    '*************************************************
    '* 実行中ファイル削除
    '*************************************************
    If bOUTPUT_EXECFILE Then
        On Error Resume Next
        objFSO.DeleteFile sRunFilePath, True
        On Error Goto 0
    End If

    '*************************************************
    '* シャットダウン実行
    '*************************************************
    objWshShell.Run "shutdown /s /f /t 0", 0, False
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

'-------------------------------------------------------------------------------
'- 対象日時算出（入力時刻が現在より前なら翌日とする）
'-------------------------------------------------------------------------------
Private Function CalcTargetDateTime( ByVal sTime ) '{{{
    Dim aParts
    aParts = Split(sTime, ":")

    Dim dTarget
    dTarget = DateSerial(Year(Now()), Month(Now()), Day(Now())) + TimeSerial(CInt(aParts(0)), CInt(aParts(1)), 0)
    If dTarget <= Now() Then
        dTarget = DateAdd("d", 1, dTarget)
    End If

    CalcTargetDateTime = dTarget
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

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}
