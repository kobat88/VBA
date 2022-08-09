Attribute VB_Name = "Main"
'Option Explicit

Public Sub main()

    Const SRC_WS_NAME = "KO (3)"                 '処理対象シート名
    Const TITLE_ROW = 7                          'タイトル行
    Const COL_LIMIT = 26                         '列の拡張上限（Z列）
    Const FIRST_COL = 1                          '開始列
    Const SAVE_DIR = "C:\Users\kobat88\Desktop\VBA\ExportToCSV" 'CSVを保存するディレクトリ
    
    Dim lastCol As Long
    Dim objUtil As Util
    Dim resArry() As Variant
    Dim i As Long
    Dim objFSO As Object
    Dim csvFileName As String, saveFilePath As String
    
    Debug.Print "Main開始"
    
    On Err GoTo ErrHandler
    
    Set objUtil = New Util
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'どの列まで項目があるか調べる
    lastCol = ThisWorkbook.Worksheets(SRC_WS_NAME).Cells(TITLE_ROW, COL_LIMIT).End(xlToLeft).Column

    '対象セル範囲（タイトル行を除く）のデータを配列に格納
    resArry = objUtil.LoadDataFromSheet(ThisWorkbook, SRC_WS_NAME, TITLE_ROW + 1, FIRST_COL, lastCol)
        
        
    '各項目を処理
    
    '図番をゼロ埋めで3桁にする
    For i = LBound(resArry, 1) To UBound(resArry, 1)
        resArry(i, 0) = Format(resArry(i, 0), "000")
    Next i
    
    '部品名から改行を除去し、""で囲む
    For i = LBound(resArry, 1) To UBound(resArry, 1)
        resArry(i, 1) = Replace(resArry(i, 1), vbLf, "")
        resArry(i, 1) = """" & resArry(i, 1) & """"
    Next i
    
    
    '配列をCSVファイルへ出力して保存
    csvFileName = objFSO.GetBaseName(ThisWorkbook.Name) & "_clean.csv"
    saveFilePath = SAVE_DIR & "\" & csvFileName
    Call objUtil.ExportArryToCSV(saveFilePath, resArry)
    
    Debug.Print "Main終了"
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number, Err.Source, Err.Description

End Sub

