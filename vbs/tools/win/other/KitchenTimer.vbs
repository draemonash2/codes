Option Explicit

'==============================================================================
'【説明】
'	指定した時間(分)の経過を待って、メッセージを表示する
'
'【使用方法】
'	1) KitchenTimer.vbs を実行して、待ち時間を入力する
'
'【覚え書き】
'	なし
'
'【改訂履歴】
'	1.0.0	2019/08/03	新規作成
'	1.1.0	2019/09/26	複数起動対応
'	1.1.0	2020/02/09	sleep化
'==============================================================================

'==============================================================================
' 設定
'==============================================================================
Const PROG_NAME = "キッチンタイマー"

Dim lWaitMinites
lWaitMinites = InputBox( "待ち時間(分)を入力してください", PROG_NAME, 1 )

If lWaitMinites = 0 Then
	MsgBox "キャンセルしました", vbYes, PROG_NAME
Else
	Dim vAnswer
	vAnswer = MsgBox( lWaitMinites & "分間のタイマーを設定しました", vbOkCancel, PROG_NAME )
	If vAnswer <> vbOk Then
		MsgBox "キャンセルしました", vbYes, PROG_NAME
	Else
		WScript.sleep(lWaitMinites * 60 * 1000)
		MsgBox lWaitMinites & "分が経過しました", vbYes, lWaitMinites & "分経過"
	End If
End If

