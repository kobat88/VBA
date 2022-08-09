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
'* Description  :渡された配列データからグラフを描画する
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
    
    'データ個数を取得
    posCntsLen = UBound(posCnts) + 1
    
    ReDim xData(posCntsLen) As Date
    ReDim yData(posCntsLen) As Long
    
    'このシート内の全てのコンテナを削除
    For Each chObj In ChartObjects
        chObj.Delete
    Next
    
    'コンテナをセット
    Set myChartObj = ChartObjects.Add(120, 10, 400, 200) 'Left,Top,Width,Height
    
    'グラフをセット
    Set myChart = myChartObj.Chart
    
    'グラフのプロパティを設定
    With myChart
        'グラフの種類
        .ChartType = xlColumnClustered           '棒グラフ
        '.ChartType = xlXYScatter                 '散布図
        
        'グラフのタイトル
        .HasTitle = True                         'タイトルあり
        .ChartTitle.Text = "PCR検査陽性者数（全国）"
        
        '凡例
        .HasLegend = False                       '凡例非表示
        '.Name = "PCR検査陽性者数"
    End With
    
    'データ系列をセット
    Set mySeries = myChart.SeriesCollection.NewSeries
    
    'データ系列にデータをセット
    With mySeries
        For i = 0 To UBound(posCnts)
            xData(i) = CDate(posCnts(i, 0))
            yData(i) = posCnts(i, 1)
        Next i
        .XValues = xData
        .Values = yData
    End With
    
    '軸のプロパティを設定
    'TODO: 下記問題の対応
    '棒グラフ、折れ線グラフだと下記はエラー。縦軸（xlValues）なら成功する。横軸はExcel上で手動でも変更できない。
    '散布図だと成功する。
    '        With myChart.Axes(xlCategory)
    '            .MajorUnit = 30                             '30項目（=30日）ごとに項目ラベル表示
    '            .TickLabels.NumberFormatLocal = "yy/mm"
    '        End With
    
End Sub

