Attribute VB_Name = "CreateGraphs"
Sub CreateGraphs()
	Dim strTableName As String
	Dim intMaxSheetNo As Integer
		For intSheetCnt As Integer To 
			' �O���t�쐬���W���[��
			CreateGraph(strTableName, intMaxSheetNo)
End Sub
