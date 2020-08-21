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
'==============================================================================

'==============================================================================
' 設定
'==============================================================================
Const sPROG_NAME = "キッチンタイマー"

Dim dWaitMinites
dWaitMinites = InputBox( "待ち時間(分)を入力してください", sPROG_NAME, 1 )
If IsEmpty(dWaitMinites) = True Then
	MsgBox "キャンセルしました", vbYes, sPROG_NAME
	WScript.Quit
ElseIf dWaitMinites = 0 Then
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

WScript.sleep(dWaitMinites * 60 * 1000) 'x[min] * 60[s] * 1000[ms]
MsgBox sPROG_NAME & vbNewLine & dWaitTime & sWaitTimeUnit & "が経過しました", vbYes, dWaitTime & sWaitTimeUnit & "経過"

