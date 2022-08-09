VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************
'* Description  :指定したブック、シート、範囲のデータを配列に格納して返す
'****************************************************************************
Public Function LoadDataFromSheet(srcWB As Workbook, srcWsName As String, firstRow As Long, firstCol As Long, lastCol As Long) As Variant

    Dim lastRow As Long
    Dim preArry() As Variant, resArry() As Variant
    Dim i As Long, j As Long
    
    On Err GoTo ErrHandler
    
    If firstRow <= 0 Or firstCol <= 0 Or firstCol > lastCol Then
        Err.Raise Number:=501, Description:="開始行、開始列、終了列の指定が不正です"
    End If
    
    
    With srcWB.Worksheets(srcWsName)
    
        'フィルタが設定されていれば解除（オートフィルタの設定は残す）
        If .FilterMode Then
            .ShowAllData
        End If
        
        '最終行取得
        lastRow = .Cells(Rows.Count, firstCol).End(xlUp).Row
        
        
        If lastRow < firstRow Then
        
            '最終行が開始行より小さければ戻り値に空の配列を設定
            resArry = VBA.Array()
        
        Else
        
            '指定のセル範囲のデータをプレ配列に格納
            preArry = Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
            
            'セル範囲を格納した配列は何故か添え字が1始まりになるので、0始まりの配列を別途定義
            ReDim resArry(0 To lastRow - firstRow, 0 To lastCol - firstCol)
            
            'プレ配列のデータを別途定義した配列に格納
            For i = 1 To UBound(preArry, 1)
                For j = 1 To UBound(preArry, 2)
                    resArry(i - 1, j - 1) = preArry(i, j)
                Next j
            Next i
            
        End If
        
    End With
    
    LoadDataFromSheet = resArry
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'****************************************************************************
'* Description  :指定したブック、シート、範囲に配列のデータを出力する
'****************************************************************************
Public Function OutDataToSheet(tgtWB As Workbook, tgtWsName As String, firstRow As Long, firstCol As Long, srcArry() As Variant)

    Dim lastRow As Long, lastCol As Long
    
    On Err GoTo ErrHandler
    
    If firstRow <= 0 Or firstCol <= 0 Then
        Err.Raise Number:=501, Description:="開始行、開始列の指定が不正です"
    End If
    
    '渡された配列が空でなければ処理する
    If Not IsEmptyArry(srcArry) Then
    
        With tgtWB.Worksheets(tgtWsName)
        
            'フィルタが設定されていれば解除（オートフィルタの設定は残す）
            If .FilterMode Then
                .ShowAllData
            End If
            
            lastRow = firstRow + UBound(srcArry, 1)
            lastCol = firstCol + UBound(srcArry, 2)
            
            Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol)) = srcArry
            
        End With
        
    End If
        
    Exit Function
        
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'****************************************************************************
'* Description  :配列が空かどうかを判定する（空の場合Trueを返す）
'****************************************************************************
Public Function IsEmptyArry(myArry() As Variant) As Boolean

    If 0 <= UBound(myArry, 1) Then
        IsEmptyArry = False
    Else
        IsEmptyArry = True
    End If
    
End Function

