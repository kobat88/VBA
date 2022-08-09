Attribute VB_Name = "DistinctData"
'データ64,000件で処理時間13秒

Public Sub DistinctData()

Const WS_NAME = "Sheet1"

Dim lastRow As Long
Dim i As Long

startTime = Timer
Debug.Print "処理開始"

Application.ScreenUpdating = False

With Worksheets(WS_NAME)

    '最終行数取得
    lastRow = .Cells(Rows.Count, 1).End(xlUp).Row

    '元の並び順を2列目に記録
    For i = 1 To lastRow
        .Cells(i, 2).Value = i
    Next i

    '1列目（URL）でソート
    .Range("A1").CurrentRegion.Sort Key1:=.Range("A1")

    '上から順に次の行と比べてドメインが同じなら次の行を着色
    For i = 1 To lastRow
        If ExtractDomain(.Cells(i, 1).Value) = ExtractDomain(.Cells(i + 1, 1).Value) Then
            .Cells(i + 1, 1).Interior.Color = RGB(200, 200, 200)
        End If
    Next i
    
    '元の順で再ソート
    .Range("A1").CurrentRegion.Sort Key1:=.Range("B1")
    
    '2列目の値をクリア
    .Columns(2).Clear
            
End With

Application.ScreenUpdating = True

Debug.Print "処理終了"
endTime = Timer
Debug.Print "処理時間：" & endTime - startTime & "@" & Now()

End Sub


Public Function ExtractDomain(ByVal urlStr As String) As String

Dim idx1 As Long, idx2 As Long
Dim str1 As String, str2 As String

idx1 = InStr(urlStr, "://")
str1 = Mid(urlStr, idx1 + 3)
idx2 = InStr(str1, "/")

If idx2 <> 0 Then
    str2 = Left(str1, idx2 - 1)
Else
    str2 = str1
End If

ExtractDomain = str2

End Function
