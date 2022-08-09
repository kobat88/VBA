Attribute VB_Name = "DeleteData2"
'Option Explicit

'Excel上でソート -> Excel上でループし検索してフラグを立てる -> フラグでフィルタして削除
'シート１：1000件 シート２：8000件で13秒

Public Sub DeleteData2()

    Const WS_NAME1 = "Sheet1"
    Const WS_NAME2 = "Sheet2"
    Const WS_NAME1_SORT = "Sheet1_sort"
    Const WS_NAME2_SORT = "Sheet2_sort"

    Dim objUtil As Util
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long, csrJ As Long, k As Long
    Dim startTime As Single, endTime As Single, procTime As Single
    Dim startTime1 As Single, endTime1 As Single
    Dim startTime2 As Single, endTime2 As Single
    Dim startTime3 As Single, endTime3 As Single
    Dim startTime4 As Single, endTime4 As Single

    startTime = Timer
    
    Debug.Print "処理開始"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set objUtil = New Util
    
    'シート１，シート２の最終行を取得
    lastRow1 = Worksheets(WS_NAME1).Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = Worksheets(WS_NAME2).Cells(Rows.Count, 1).End(xlUp).Row
    
    
    startTime1 = Timer
    'シート１データをソート用シート１にコピーしてソート
    Worksheets(WS_NAME1).Copy After:=Worksheets(WS_NAME1)
    ActiveSheet.Name = WS_NAME1_SORT
    
    With Worksheets(WS_NAME1_SORT)
        '元の並び順を9列目に記録
        For k = 2 To lastRow1
            .Cells(k, 9).Value = k
        Next
        '1列目でソート
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), Header:=xlYes
    End With
    
    'シート２データをソート用シート２にコピーしてソート
    Worksheets(WS_NAME2).Copy After:=Worksheets(WS_NAME2)
    ActiveSheet.Name = WS_NAME2_SORT
    
    With Worksheets(WS_NAME2_SORT)
        '元の並び順を9列目に記録
        For k = 2 To lastRow2
            .Cells(k, 9).Value = k
        Next
        '1列目でソート
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), Header:=xlYes
    End With
    endTime1 = Timer
    Debug.Print "ソート処理:" & endTime1 - startTime1
    
    
    'メイン処理
    startTime3 = Timer1
    csrJ = 2
    'ソート用シート２を前から順に処理
    For i = 2 To lastRow2
        '10列目に"not exist"を格納
        Worksheets(WS_NAME2_SORT).Cells(i, 10) = "not exist"
        'ソート用シート１を前から順にチェック
        For j = csrJ To lastRow1
            'ソート用シート２のA列=ソート用シート１のA列であれば、ソート用シート２の10列目を"exist"に更新し、ソート用シート２の次の行へ
            If Worksheets(WS_NAME2_SORT).Cells(i, 1) = Worksheets(WS_NAME1_SORT).Cells(j, 1) Then
                Worksheets(WS_NAME2_SORT).Cells(i, 10) = "exist"
                csrJ = j
                Exit For
            End If
        Next j
    Next i
    endTime3 = Timer
    Debug.Print "メイン処理:" & endTime3 - startTime3

    
    'ソート用シート２を10列目="exist"でフィルター
    Worksheets(WS_NAME2_SORT).Range("A1").AutoFilter 10, "exist"
    
    'フィルターした行を削除
    With Worksheets(WS_NAME2_SORT).Range("A1").CurrentRegion
        .Resize(.Rows.Count - 1).Offset(1, 0).Delete
    End With
    
    'フィルター解除
    Worksheets(WS_NAME2_SORT).Range("A1").AutoFilter
    
    'ソート用シート２のデータを元の並び順でソート
    Worksheets(WS_NAME2_SORT).Range("A1").CurrentRegion.Sort Key1:=Worksheets(WS_NAME2_SORT).Range("I1"), Header:=xlYes

    'ソート用シート２の9､10列目をクリア
    Worksheets(WS_NAME2_SORT).Columns(9).Clear
    Worksheets(WS_NAME2_SORT).Columns(10).Clear

    'シート２を削除
    Worksheets(WS_NAME2).Delete

    'ソート用シート２をシート２へコピー
    Worksheets(WS_NAME2_SORT).Copy After:=Worksheets(WS_NAME1)
    ActiveSheet.Name = WS_NAME2

    'ソート用シートを削除
    Worksheets(WS_NAME1_SORT).Delete
    Worksheets(WS_NAME2_SORT).Delete
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = False

    Debug.Print "処理終了"
    
    endTime = Timer
    procTime = endTime - startTime
    
    Debug.Print "処理時間:" & procTime

End Sub

