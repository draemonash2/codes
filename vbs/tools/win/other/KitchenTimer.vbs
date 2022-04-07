Option Explicit

'==============================================================================
'【説明】
'	指定した時間(分)の経過を待って、メッセージを表示する
'
'【使用方法】
'	1) KitchenTimer.vbs を実行して、待ち時間(分)を入力する
'
'【覚え書き】
'	なし
'
'【改訂履歴】
'	1.0.0	2019/08/03	新規作成
'	1.1.0	2019/09/26	複数起動対応
'	1.1.1	2020/02/09	sleep化
'	1.1.2	2020/08/21	秒/時間表示対応
'	1.2.0	2021/01/14	タイトル出力機能追加
'	1.2.1	2021/01/30	デフォルトタイトル修正
'	1.2.2	2022/04/07	実行中ファイル出力/削除処理追加
'==============================================================================

'==============================================================================
'= インクルード
'==============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ConvDate2String()

'==============================================================================
' 設定
'==============================================================================
Const sPROG_NAME = "キッチンタイマー"

'==============================================================================
'= 本処理
'==============================================================================
'*************************************************
'* タイマー設定
'*************************************************
Dim dWaitMinites
dWaitMinites = InputBox( "待ち時間(分)を入力してください", sPROG_NAME, 1 )
If IsEmpty(dWaitMinites) = True Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
ElseIf dWaitMinites = 0 Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
End If

Dim sOutputMsg
sOutputMsg = InputBox( "タイトルを入力してください", sPROG_NAME )
If IsEmpty(sOutputMsg) = True Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
End If

Dim dWaitTime
Dim sWaitTimeUnit
If dWaitMinites < 1 Then
	dWaitTime = Round( dWaitMinites * 60, 2 )
	sWaitTimeUnit = "秒"
ElseIf dWaitMinites >= 60 Then
	dWaitTime = Round( dWaitMinites / 60, 2 )
	sWaitTimeUnit = "時間"
Else
	dWaitTime = Round( dWaitMinites, 2 )
	sWaitTimeUnit = "分"
End IF

Dim vAnswer
vAnswer = MsgBox( dWaitTime & sWaitTimeUnit & "のタイマーを設定しました", vbOkCancel, sPROG_NAME )
If IsEmpty(vAnswer) = True Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
ElseIf vAnswer <> vbOk Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
End If

'*************************************************
'* 実行中ファイル出力
'*************************************************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sStartTime
sStartTime = ConvDate2String(Now(), 0)
Dim sTrgtFilePath
sTrgtFilePath = objWshShell.SpecialFolders("Desktop") & "\running_kitchentimer_" & dWaitMinites & "min[" & sOutputMsg & "]_" & sStartTime & ".txt"
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objTxtFile
On Error Resume Next
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
objTxtFile.WriteLine "running..."
objTxtFile.Close
On Error Goto 0

'*************************************************
'* 待ち処理
'*************************************************
WScript.sleep(dWaitMinites * 60 * 1000) 'x[min] * 60[s] * 1000[ms]

'*************************************************
'* メッセージ出力
'*************************************************
If sOutputMsg = "" Then
	MsgBox sPROG_NAME & vbNewLine & dWaitTime & sWaitTimeUnit & "が経過しました", vbYes, dWaitTime & sWaitTimeUnit & "経過"
Else
	MsgBox sPROG_NAME & vbNewLine & dWaitTime & sWaitTimeUnit & "が経過しました", vbYes, sOutputMsg
End If

'*************************************************
'* 実行中ファイル削除
'*************************************************
On Error Resume Next
objFSO.DeleteFile sTrgtFilePath, True
On Error Goto 0

'==============================================================================
'= インクルード関数
'==============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
