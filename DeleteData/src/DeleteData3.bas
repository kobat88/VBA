Attribute VB_Name = "DeleteData3"
'Option Explicit

'Excel上でシート１だけソート -> Excel上でmatchで検索してフラグを立てる -> フラグでフィルタして削除
'シート１：1000件 シート２：8000件で5秒

Public Sub DeleteData3()

    Const WS_NAME1 = "Sheet1"
    Const WS_NAME2 = "Sheet2"
    Const WS_NAME1_SORT = "Sheet1_sort"

    Dim objUtil As Util
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, k As Long
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
    
    
    'メイン処理
    startTime3 = Timer1
    'シート２を前から順に処理
    For i = 2 To lastRow2
        '9列目に"not exist"を格納
        Worksheets(WS_NAME2).Cells(i, 9) = "not exist"
        'シート２のA列の値で、ソート用シート１のA列をmatch検索（見つからなければエラーが返る）
        matchIdx = Application.Match(Worksheets(WS_NAME2).Cells(i, 1), Worksheets(WS_NAME1_SORT).Range("A:A"), 0)
        '検索結果が見つかれば、シート２の10列目を"exist"で更新
        If Not IsError(matchIdx) Then
            Worksheets(WS_NAME2).Cells(i, 9) = "exist"
        End If
    Next i
    endTime3 = Timer
    Debug.Print "メイン処理:" & endTime3 - startTime3

    
    'シート２を10列目="exist"でフィルター
    Worksheets(WS_NAME2).Range("A1").AutoFilter 9, "exist"

    'フィルターした行を削除
    With Worksheets(WS_NAME2).Range("A1").CurrentRegion
        .Resize(.Rows.Count - 1).Offset(1, 0).Delete
    End With

    'フィルター解除
    Worksheets(WS_NAME2).Range("A1").AutoFilter

    'シート２の9列目をクリア
    Worksheets(WS_NAME2).Columns(9).Clear

    'ソート用シート１を削除
    Worksheets(WS_NAME1_SORT).Delete

    Application.ScreenUpdating = True
    Application.DisplayAlerts = False

    Debug.Print "処理終了"
    
    endTime = Timer
    procTime = endTime - startTime
    
    Debug.Print "処理時間:" & procTime & "@" & Now()

End Sub

