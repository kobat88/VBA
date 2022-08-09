Attribute VB_Name = "CreateManyData"
'*************************************************************
'* Description: シートにある複数行データをコピーして倍の行にする
'*************************************************************
Public Sub CreateManyData()

Const WS_NAME = "Sheet1"
Const FIRST_ROW = 1

Dim lastRow As Long

With Worksheets(WS_NAME)

lastRow = .Cells(Rows.Count, 1).End(xlUp).Row

.Range(Rows(FIRST_ROW), .Rows(lastRow)).Copy Destination:=.Rows(lastRow + 1)

End With

End Sub
