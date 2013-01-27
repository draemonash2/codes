Attribute VB_Name = "CreateGraphs"
Sub CreateGraphs()
	Dim strTableName As String
	Dim intMaxSheetNo As Integer
		For intSheetCnt As Integer To 
			' グラフ作成モジュール
			CreateGraph(strTableName, intMaxSheetNo)
End Sub
