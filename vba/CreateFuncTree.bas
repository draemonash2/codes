Attribute VB_Name = "CreateFuncTree"
Sub CreateFuncTree()
    ' �ϐ���`
    Dim intHeight As Integer                ' �ǉ�����e�L�X�g�{�b�N�X�̈ʒu(����)�
    Dim intStartRow As Integer              ' �X�^�[�g����s��
    Dim intConnecterBeginPointY As Integer  ' �R�l�N�^�n�_�̐����ʒu
    Dim intConnecterBeginPointX As Integer  ' �R�l�N�^�n�_�̐����ʒu
    Dim shpObjectBox As Shape               ' �֐����{�b�N�X��`
    Dim shpObjectLine As Shape              ' �R�l�N�^��`
    
    ' �֐����{�b�N�X�����ʒu��`
    intHeight = 25
    intStartRow = 2
    
    For intSearchRow = intStartRow To (Range("B2").End(xlDown).Row)
        ' === �֐����{�b�N�X���� ===
            ' �I�u�W�F�N�g����
            Set shpObjectBox = ActiveSheet.Shapes.AddShape( _
                Type:=msoShapeFlowchartPredefinedProcess, _
                Left:=200, _
                Top:=intHeight * (intSearchRow - 1), _
                Width:=100, _
                Height:=100)
            
            ' �����ݒ�
            shpObjectBox.Fill.ForeColor.RGB = RGB(128, 0, 0)     ' �w�i�F
            shpObjectBox.Line.ForeColor.RGB = RGB(0, 0, 0)       ' ���̐F
            shpObjectBox.Line.Weight = 2                         ' ���̑���
            shpObjectBox.Select
            Selection.Characters.Text = Cells(intSearchRow, 2)   ' �e�L�X�g�ɐ����̓��e��ݒ�
            Selection.AutoSize = True                            ' �����T�C�Y�����ɂ���
        
        ' === �R�l�N�^���� ===
            ' �I�u�W�F�N�g����
            intConnecterBeginPointY = (intHeight * (intSearchRow - 1)) - 10
            intConnecterBeginPointX = 200
            Set shpObjectLine = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, intConnecterBeginPointX, intConnecterBeginPointY, 0, 0)
            ' �����ݒ�
            shpObjectLine.Line.ForeColor.RGB = RGB(0, 0, 0)      ' ���̐F
            shpObjectLine.Line.Weight = 2                        ' ���̑���
            ' �R�l�N�^�ڑ�
            shpObjectLine.Select
            Selection.ShapeRange.ConnectorFormat.EndConnect shpObjectBox, 1
            
   Next intSearchRow
   
End Sub
