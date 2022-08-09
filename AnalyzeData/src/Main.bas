Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()

    Const DL_FILE_URL = "https://www.mhlw.go.jp/content/pcr_positive_daily.csv"
    Const SAVE_DIR = "C:\Users\kobat88\Desktop\VBA\AnalyzeData\CovidDL"

    Dim objModel As Model
    Dim objUtil As Util
    Dim filePath As String
    Dim posCnts() As Variant
    
    Debug.Print "Main開始"
    
    On Error GoTo ErrHandler

    With Application
        .ScreenUpdating = False
    End With
    
    Set objModel = New Model
    Set objUtil = New Util
    
    '日ごとの陽性者数のcsvファイルをダウンロード
    Call objUtil.DownloadFile(DL_FILE_URL, SAVE_DIR)
    
    filePath = SAVE_DIR & "\" & objUtil.GetFileNameFromURL(DL_FILE_URL)
    
    '日ごとの陽性者数の累計を求めてResultシートに出力
    Result.ProcName = "CalcPosTotal"
    Result.ProcResult = objModel.CalcPosTotal(filePath)
    Result.ProcDatetime = Now
    Result.ErrDesc = vbNullString
    Debug.Print "結果出力完了"
    
    'Graphシートにグラフ出力
    posCnts = objUtil.LoadDataFromCSV(filePath)
    Call Graph.DrawGraph(posCnts)
    
    Debug.Print "グラフ出力完了"
    
    GoTo Finally

ErrHandler:
    Result.ProcName = "CalcPosTotal"
    Result.ProcResult = 0
    Result.ProcDatetime = Now
    Result.ErrDesc = Err.Description

    Debug.Print Err.Number, Err.Source, Err.Description

    Resume Finally
    
Finally:
    With Application
        .ScreenUpdating = True
    End With
    
End Sub



Public Sub PrintSheet()

    Const PRT_WS_NAME = "Graph"

    With Worksheets(PRT_WS_NAME).PageSetup
        '用紙横向き
        .Orientation = xlLandscape
        '余白設定
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        'センタリング
        .CenterHorizontally = True
        .CenterVertically = True
        '用紙一枚に納まるように印刷
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    'PageSetupを高速化するメソッドだがあまり変わらない
    'Application.PrintCommunication = True
    
    Worksheets(PRT_WS_NAME).PrintOut Preview:=True

End Sub




