Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Public Sub Main()

    Const SRC_WS_NAME = "シート②"
    Const TGT_WS_NAME = "シート①"

    Dim objUtil As Util_OptionBase1
    Dim srcArry() As Variant, tgtArry() As Variant
    Dim i As Long, j As Long, k As Long, updateCnt As Long
    Dim startTime As Single, endTime As Single
    
    Set objUtil = New Util_OptionBase1
    
    startTime = Timer

    Debug.Print "Main開始"

    On Err GoTo ErrHandler

    Application.ScreenUpdating = False

    'コピー元シートのデータをコピー元配列に格納
    srcArry = objUtil.LoadDataFromSheet(ThisWorkbook, SRC_WS_NAME, 2, 1, 8)

    'コピー先シートの既存データをコピー先配列に格納
    tgtArry = objUtil.LoadDataFromSheet(ThisWorkbook, TGT_WS_NAME, 2, 1, 8)
    
    
    'コピー元配列が空でなければ処理する
    If Not objUtil.IsEmptyArry(srcArry) Then
    
        'コピー先配列の最終行を取得
        If objUtil.IsEmptyArry(tgtArry) Then
            k = 1
        Else
            k = UBound(tgtArry, 1) + 1
        End If
        
    
        'コピー元配列を一行ずつ処理
        For i = LBound(srcArry, 1) To UBound(srcArry, 1)
            
            'コピー元配列の案件IDでコピー先配列を検索し、ヒットしたら該当行を更新
            updateCnt = 0
            For j = LBound(tgtArry, 1) To UBound(tgtArry, 1)
                If tgtArry(j, 8) = srcArry(i, 1) Then
                    tgtArry(j, 1) = srcArry(i, 2)
                    tgtArry(j, 2) = srcArry(i, 3)
                    tgtArry(j, 3) = srcArry(i, 4)
                    tgtArry(j, 4) = srcArry(i, 5)
                    tgtArry(j, 5) = srcArry(i, 6)
                    tgtArry(j, 6) = srcArry(i, 7)
                    tgtArry(j, 7) = srcArry(i, 8)
                    updateCnt = updateCnt + 1
                End If
            Next j
            
            'コピー元配列の案件IDがコピー先配列に存在しなければ、コピー先配列の最終行の次の行に該当行を追加
            If updateCnt = 0 Then
            
                If objUtil.IsEmptyArry(tgtArry) Then
                    'コピー先配列が空ならば、1行で再定義
                    ReDim tgtArry(1, 8) As Variant
                Else
                    'コピー先配列が空でなければ、コピー先配列の行数を１増やす
                    tgtArry = objUtil.ExpandFirstDimOfArry(tgtArry, 1)
                End If
            
                tgtArry(k, 1) = srcArry(i, 2)
                tgtArry(k, 2) = srcArry(i, 3)
                tgtArry(k, 3) = srcArry(i, 4)
                tgtArry(k, 4) = srcArry(i, 5)
                tgtArry(k, 5) = srcArry(i, 6)
                tgtArry(k, 6) = srcArry(i, 7)
                tgtArry(k, 7) = srcArry(i, 8)
                tgtArry(k, 8) = srcArry(i, 1)
                k = k + 1
            End If
        Next i
        
        'コピー先配列のデータをシート①へ出力
        Call objUtil.OutDataToSheet(ThisWorkbook, TGT_WS_NAME, 2, 1, tgtArry)
    
    Else
        Debug.Print SRC_WS_NAME & "のデータが0件です"
    End If
    
    Set objUtil = Nothing
    
    Application.ScreenUpdating = True

    Debug.Print "Main終了"
    
    endTime = Timer
    Debug.Print "Main:" & Now() & "," & (endTime - startTime) & "sec"
    
    Exit Sub

ErrHandler:
    Debug.Print Err.Number, Err.Description

End Sub

'*************************************************************************************
'* Description  :二次元動的配列のデータを保持したまま一次元目を要素数をiだけ増やす
'* （多次元配列のReDim Preserveでは最終次元の要素数しか増やせないため、この関数を用意）
'*************************************************************************************
'Public Function ExpandFirstDimOfArry(ByVal srcArry As Variant, i As Long) As Variant
'    Dim tmpArry() As Variant
'
'    '一次元目と二次元目を交換
'    tmpArry = WorksheetFunction.Transpose(srcArry)
'
'    ReDim Preserve tmpArry(UBound(tmpArry, 1), UBound(tmpArry, 2) + i)
'
'    ExpandFirstDimOfArry = WorksheetFunction.Transpose(tmpArry)
'
'End Function

