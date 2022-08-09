Attribute VB_Name = "PrintToPDF"
Public Sub PrintToPDF()

    Const WS_NAME = "Sheet1"

    Dim pdfFilePath As String

    pdfFilePath = ThisWorkbook.Path & "\test1.pdf"

    With ThisWorkbook.Worksheets(WS_NAME)

        '改ページ設定
        .Rows(2).PageBreak = xlPageBreakManual
    
        '他にも印刷範囲の設定で余白など調整できる
    
        .PrintOut _
        ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, _
        PrToFileName:=pdfFilePath, _
        Preview:=False

    End With
        
End Sub


