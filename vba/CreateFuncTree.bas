Attribute VB_Name = "CreateFuncTree"
Sub CreateFuncTree()
    ' 変数定義
    Dim intHeight As Integer                ' 追加するテキストボックスの位置(高さ)基準
    Dim intStartRow As Integer              ' スタートする行数
    Dim intConnecterBeginPointY As Integer  ' コネクタ始点の垂直位置
    Dim intConnecterBeginPointX As Integer  ' コネクタ始点の水平位置
    Dim shpObjectBox As Shape               ' 関数名ボックス定義
    Dim shpObjectLine As Shape              ' コネクタ定義
    
    ' 関数名ボックス生成位置定義
    intHeight = 25
    intStartRow = 2
    
    For intSearchRow = intStartRow To (Range("B2").End(xlDown).Row)
        ' === 関数名ボックス生成 ===
            ' オブジェクト生成
            Set shpObjectBox = ActiveSheet.Shapes.AddShape( _
                Type:=msoShapeFlowchartPredefinedProcess, _
                Left:=200, _
                Top:=intHeight * (intSearchRow - 1), _
                Width:=100, _
                Height:=100)
            
            ' 書式設定
            shpObjectBox.Fill.ForeColor.RGB = RGB(128, 0, 0)     ' 背景色
            shpObjectBox.Line.ForeColor.RGB = RGB(0, 0, 0)       ' 線の色
            shpObjectBox.Line.Weight = 2                         ' 線の太さ
            shpObjectBox.Select
            Selection.Characters.Text = Cells(intSearchRow, 2)   ' テキストに数式の内容を設定
            Selection.AutoSize = True                            ' 自動サイズ調整にする
        
        ' === コネクタ生成 ===
            ' オブジェクト生成
            intConnecterBeginPointY = (intHeight * (intSearchRow - 1)) - 10
            intConnecterBeginPointX = 200
            Set shpObjectLine = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, intConnecterBeginPointX, intConnecterBeginPointY, 0, 0)
            ' 書式設定
            shpObjectLine.Line.ForeColor.RGB = RGB(0, 0, 0)      ' 線の色
            shpObjectLine.Line.Weight = 2                        ' 線の太さ
            ' コネクタ接続
            shpObjectLine.Select
            Selection.ShapeRange.ConnectorFormat.EndConnect shpObjectBox, 1
            
   Next intSearchRow
   
End Sub
