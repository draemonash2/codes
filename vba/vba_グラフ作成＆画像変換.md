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
    '=== �O���t�쐬 ===
    Set coGraphObj = .ChartObjects.Add( _
                            dGraphPosX, _
                            dGraphPosY, _
                            dGraphWidth, _
                            dGraphHeight _
                        )
    With coGraphObj.Chart
        .ChartType = xlXYScatterLines                             '�U�z�}�Ɏw��
        .SetSourceData Source:=Union(rXAxsRng, rDataRng)          '�f�[�^�͈͎w��
        .Axes(xlCategory, xlPrimary).MinimumScale = dGraphXRngMin '�w���ŏ��l
        .Axes(xlCategory, xlPrimary).MaximumScale = dGraphXRngMax '�w���ő�l
        .Axes(xlCategory, xlPrimary).TickLabels.Orientation = -90 '�w�����x���p�x
        .Axes(xlCategory, xlPrimary).TickLabelPosition = xlLow    '�w�����x���ʒu
        .PlotArea.Select                                          '�v���b�g�G���A�I���iPlotArea �I�u�W�F�N�g�̓O���t���A�N�e�B�u�ɂ��Ȃ��Ƒ���ł��Ȃ��j
        .PlotArea.InsideLeft = 30                                 '�v���b�g�G���A�̉��ʒu
        .PlotArea.InsideHeight = dGraphHeight - 100               '�v���b�g�G���A�̍���
        .PlotArea.InsideWidth = dGraphWidth - 70                  '�v���b�g�G���A�̕�
        .HasLegend = False                                        '�}��Ȃ�
    End With
 
    '=== �O���t�ˉ摜�ϊ� ===
    coGraphObj.Cut
    .Cells(lGraphOutpRow, lGraphOutpClm).Select
    .PasteSpecial Format:="�} (JPEG)"
 
    Set coGraphObj = Nothing
End With
```
