Attribute VB_Name = "CombineData"
Public Sub Main()

    Dim srcArry1() As Variant, srcArry2() As Variant, resArry() As Variant
    Dim i As Long, j As Long, k As Long, combiLen As Long
    Dim objUtil As Util
    
    Debug.Print "Main開始"
    
    Set objUtil = New Util
    
    '組み合わせるデータを列ごとに配列に格納
    srcArry1 = objUtil.LoadDataFromSheet(ThisWorkbook, "Sheet1", 2, 1, 1)
    srcArry2 = objUtil.LoadDataFromSheet(ThisWorkbook, "Sheet1", 2, 2, 2)
    
    '組み合わせるデータのいずれかの列が空の場合エラー
    If objUtil.IsEmptyArry(srcArry1) Or objUtil.IsEmptyArry(srcArry2) Then
        Debug.Print "エラー：データ列のいずれかが空です"
        Exit Sub
    End If
    
    '組み合わせの数を取得
    combiLen = (UBound(srcArry1) + 1) * (UBound(srcArry2) + 1)
    
    '行数＝組み合わせの数で結果を格納する配列を再定義
    ReDim resArry(combiLen - 1, 0) As Variant

    '組み合わせを生成して結果配列に格納
    k = 0
    For i = LBound(srcArry1, 1) To UBound(srcArry1, 1)
        For j = LBound(srcArry2, 1) To UBound(srcArry2, 1)
            resArry(k, 0) = srcArry2(j, 0) & "の" & srcArry1(i, 0)
            k = k + 1
        Next j
    Next i

    '結果配列をシートに出力
    Call objUtil.OutDataToSheet(ThisWorkbook, "sheet1", 2, 3, resArry)

    Debug.Print "Main終了"

End Sub

