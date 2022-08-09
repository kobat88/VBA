Attribute VB_Name = "Main"
Option Explicit

'*************************************************************************************************
'* Description  :SRC_DIR下の全エクセルファイルの全シートのデータをTGT_SHEET_NAMEシートへ出力する
'*************************************************************************************************
Public Sub Main()

    Const SRC_DIR = "C:\Users\kobat88\Desktop\VBA\MergeData\TestDir\" '末尾は"\"までつける
    Const TGT_SHEET_NAME = "Merge"
    Const TITLE_ROW_NUM = 2
    Const FILE_NAME_COL_NUM = 2
    Const SHEET_NAME_COL_NUM = 3
    Const DATA_COL_NUM = 4
    
    Dim objUtil As Util
    Dim srcFileName As String
    Dim srcWB As Workbook
    Dim srcWS As Worksheet
    Dim resArry() As Variant
    Dim i As Long
    
    Debug.Print "Main開始"
    
    Application.ScreenUpdating = False
    
    With ThisWorkbook.Sheets(TGT_SHEET_NAME)
        
        'タイトル行より下の行を削除
        .Rows(TITLE_ROW_NUM & ":" & Cells.Rows.Count).Delete
        
        Set objUtil = New Util
        
        On Err GoTo ErrHandler
        
        'SRC_DIR下にあるエクセルファイルのファイル名を取得
        srcFileName = Dir(SRC_DIR & "*.xls?")
        
        i = TITLE_ROW_NUM + 1
        
        'コピー元ファイルを１ファイルずつ開いて処理
        Do While srcFileName <> ""
            Set srcWB = Workbooks.Open(fileName:=SRC_DIR & srcFileName, ReadOnly:=True, UpdateLinks:=0)

            'ファイル名出力
            .Cells(i, FILE_NAME_COL_NUM).Value = srcWB.Name
            
            'コピー元ファイルの全シートを順番に処理
            For Each srcWS In srcWB.Worksheets
            
                'シート名出力
                .Cells(i, SHEET_NAME_COL_NUM).Value = srcWS.Name
                
                '対象シートのデータを取得
                resArry = objUtil.LoadDataFromSheet(srcWB, srcWS.Name, 2, 2, 3)
                
                '取得データ出力
                Range(.Cells(i, DATA_COL_NUM), .Cells(i + UBound(resArry, 1), DATA_COL_NUM + UBound(resArry, 2))) = resArry
                
                i = i + UBound(resArry, 1)
            Next srcWS
            
            srcWB.Close
            
            '次のファイル名を取得
            srcFileName = Dir()
        Loop
    
    End With
    
    Application.ScreenUpdating = True
    
    Debug.Print "Main終了"
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

