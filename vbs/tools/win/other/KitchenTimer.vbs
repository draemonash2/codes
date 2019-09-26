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
'==============================================================================

'==============================================================================
' 設定
'==============================================================================
Const PROG_NAME = "キッチンタイマー"

Dim lWaitMinites
lWaitMinites = InputBox( "待ち時間(分)を入力してください", PROG_NAME, 1 )

If lWaitMinites = 0 Then
	MsgBox _
		"キャンセルしました", _
		vbYes, _
		PROG_NAME
Else
	Dim vAnswer
	vAnswer = MsgBox( _
		lWaitMinites & "分間のタイマーを設定しました", _
		vbOkCancel, _
		PROG_NAME _
	)
	If vAnswer <> vbOk Then
		MsgBox _
			"キャンセルしました", _
			vbYes, _
			PROG_NAME
	Else
		Dim objWsh
		Set objWsh = WScript.CreateObject("WScript.Shell")
		objWsh.Run "KitchenTimerPost.vbs "& lWaitMinites
	End If
End If

