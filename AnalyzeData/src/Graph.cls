VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'********************************************************************************************
'* Description  :�n���ꂽ�z��f�[�^����O���t��`�悷��
'********************************************************************************************
Public Sub DrawGraph(posCnts() As Variant)

    Dim myChartObj As ChartObject
    Dim chObj As ChartObject

    Dim myChart As Chart
    Dim mySeries As Series
    Dim xData() As Date
    Dim yData() As Long
    Dim posCntsLen As Long
    Dim i As Long
    
    '�f�[�^�����擾
    posCntsLen = UBound(posCnts) + 1
    
    ReDim xData(posCntsLen) As Date
    ReDim yData(posCntsLen) As Long
    
    '���̃V�[�g���̑S�ẴR���e�i���폜
    For Each chObj In ChartObjects
        chObj.Delete
    Next
    
    '�R���e�i���Z�b�g
    Set myChartObj = ChartObjects.Add(120, 10, 400, 200) 'Left,Top,Width,Height
    
    '�O���t���Z�b�g
    Set myChart = myChartObj.Chart
    
    '�O���t�̃v���p�e�B��ݒ�
    With myChart
        '�O���t�̎��
        .ChartType = xlColumnClustered           '�_�O���t
        '.ChartType = xlXYScatter                 '�U�z�}
        
        '�O���t�̃^�C�g��
        .HasTitle = True                         '�^�C�g������
        .ChartTitle.Text = "PCR�����z���Ґ��i�S���j"
        
        '�}��
        .HasLegend = False                       '�}���\��
        '.Name = "PCR�����z���Ґ�"
    End With
    
    '�f�[�^�n����Z�b�g
    Set mySeries = myChart.SeriesCollection.NewSeries
    
    '�f�[�^�n��Ƀf�[�^���Z�b�g
    With mySeries
        For i = 0 To UBound(posCnts)
            xData(i) = CDate(posCnts(i, 0))
            yData(i) = posCnts(i, 1)
        Next i
        .XValues = xData
        .Values = yData
    End With
    
    '���̃v���p�e�B��ݒ�
    'TODO: ���L���̑Ή�
    '�_�O���t�A�܂���O���t���Ɖ��L�̓G���[�B�c���ixlValues�j�Ȃ琬������B������Excel��Ŏ蓮�ł��ύX�ł��Ȃ��B
    '�U�z�}���Ɛ�������B
    '        With myChart.Axes(xlCategory)
    '            .MajorUnit = 30                             '30���ځi=30���j���Ƃɍ��ڃ��x���\��
    '            .TickLabels.NumberFormatLocal = "yy/mm"
    '        End With
    
End Sub
