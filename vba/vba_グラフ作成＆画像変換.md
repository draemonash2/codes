```vba
With shTrgtSht
    lStrtRow = GRAPH_TRGT_STRT_ROW
    lEndRow  = GRAPH_TRGT_END_ROW
    lGraphOutpRow = GRAPH_OUTP_ROW
    lGraphOutpClm = GRAPH_OUTP_CLM
    dGraphPosY = .Cells(lGraphOutpRow, lGraphOutpClm).Top
    dGraphPosX = .Cells(lGraphOutpRow, lGraphOutpClm).Left
    dCellWidth = ( _
                        .Cells(lGraphOutpRow, lGraphOutpClm + 1).Left - _
                        .Cells(lGraphOutpRow, lGraphOutpClm).Left _
                    )
    dCellHeight = ( _
                        .Cells(lGraphOutpRow + 1, lGraphOutpClm).Top - _
                        .Cells(lGraphOutpRow, lGraphOutpClm).Top _
                    )
    dGraphWidth = dCellWidth * GRAPH_WIDTH_CLM_NUM
    dGraphHeight = dCellHeight * GRAPH_HEIGHT_ROW_NUM
    dGraphXRngMin = gtBasicInfo.tCycInfo.dTimeFixMin
    dGraphXRngMax = gtBasicInfo.tCycInfo.dTimeFixMax
    Set rXAxsRng = .Range( _
                        .Cells(lStrtRow, GRAPH_XAXIS_CLM), _
                        .Cells(lEndRow, GRAPH_XAXIS_CLM) _
                    )
    Set rDataRng = .Range( _
                        .Cells(lStrtRow, GRAPH_DATA_CLM), _
                        .Cells(lEndRow, GRAPH_DATA_CLM) _
                    )
    '=== グラフ作成 ===
    Set coGraphObj = .ChartObjects.Add( _
                            dGraphPosX, _
                            dGraphPosY, _
                            dGraphWidth, _
                            dGraphHeight _
                        )
    With coGraphObj.Chart
        .ChartType = xlXYScatterLines                             '散布図に指定
        .SetSourceData Source:=Union(rXAxsRng, rDataRng)          'データ範囲指定
        .Axes(xlCategory, xlPrimary).MinimumScale = dGraphXRngMin 'Ｘ軸最小値
        .Axes(xlCategory, xlPrimary).MaximumScale = dGraphXRngMax 'Ｘ軸最大値
        .Axes(xlCategory, xlPrimary).TickLabels.Orientation = -90 'Ｘ軸ラベル角度
        .Axes(xlCategory, xlPrimary).TickLabelPosition = xlLow    'Ｘ軸ラベル位置
        .PlotArea.Select                                          'プロットエリア選択（PlotArea オブジェクトはグラフをアクティブにしないと操作できない）
        .PlotArea.InsideLeft = 30                                 'プロットエリアの横位置
        .PlotArea.InsideHeight = dGraphHeight - 100               'プロットエリアの高さ
        .PlotArea.InsideWidth = dGraphWidth - 70                  'プロットエリアの幅
        .HasLegend = False                                        '凡例なし
    End With
 
    '=== グラフ⇒画像変換 ===
    coGraphObj.Cut
    .Cells(lGraphOutpRow, lGraphOutpClm).Select
    .PasteSpecial Format:="図 (JPEG)"
 
    Set coGraphObj = Nothing
End With
```
