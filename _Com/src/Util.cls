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

'********************************************************************************************
'* Windows API使用宣言
'********************************************************************************************

'ごみ箱へ移動するAPI
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" _
(lpFileOp As SHFILEOPSTRUCT) As Long
'SHFileOperation関数に渡すユーザー定義型
Private Type SHFILEOPSTRUCT
    hwnd As Long                                 'ウィンドウハンドル
    wFunc As Long                                '実行する操作
    pFrom As String                              '対象ファイル名
    pTo As String                                '目的ファイル名
    fFlags As Integer                            'フラグ
    fAnyOperationsAborted As Long                '結果
    hNameMappings As Long                        'ファイル名マッピングオブジェクト
    lpszProgressTitle As String                  'ダイアログのタイトル
End Type

Private Const FO_DELETE = &H3                    '削除する
Private Const FOF_ALLOWUNDO = &H40               'ごみ箱に送る

'URLからファイルダウンロードするAPI
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long


'****************************************************************************
'* 以下にUtilのFunctionを定義
'****************************************************************************

'****************************************************************************************
'* Description  :指定したブック、シート、範囲のデータを配列に格納して返す（Option Base 0 用）
'* 当Functionは、Option Base 0 を指定した（省略可）モジュールで使用する
'****************************************************************************************
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
            
            'セル範囲を格納した配列は仕様上添え字が1始まりになるので、0始まりの配列を別途定義
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

'*************************************************************************************
'* Description  :指定したブック、シート、範囲に配列のデータを出力する（Option Base 0 用）
'* 当Functionは、Option Base 0（省略可） を指定したモジュールで使用する
'*************************************************************************************
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

'****************************************************************************************
'* Description  :指定したブック、シート、範囲のデータを配列に格納して返す（Option Base 1 用）
'* 当Functionは、Option Base 1 を指定したモジュールで使用する
'****************************************************************************************
Public Function LoadDataFromSheetBase1(srcWB As Workbook, srcWsName As String, firstRow As Long, firstCol As Long, lastCol As Long) As Variant

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

            '指定のセル範囲のデータを配列に格納
            resArry = Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
            
        End If
        
    End With
    
    LoadDataFromSheetBase1 = resArry
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'************************************************************************************
'* Description  :指定したブック、シート、範囲に配列のデータを出力する（Option Base 1 用）
'* 当Functionは、Option Base 1 を指定したモジュールで使用する
'************************************************************************************
Public Function OutDataToSheetBase1(tgtWB As Workbook, tgtWsName As String, firstRow As Long, firstCol As Long, srcArry() As Variant)

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
            
            lastRow = firstRow + UBound(srcArry, 1) - 1
            lastCol = firstCol + UBound(srcArry, 2) - 1
            
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

'*************************************************************************************
'* Description  :二次元動的配列のデータを保持したまま一次元目を要素数をiだけ増やす
'* （多次元配列のReDim Preserveでは最終次元の要素数しか増やせないため、この関数を用意）
'* Transposeの仕様上、Option Base 1とする必要あり
'*************************************************************************************
Public Function ExpandFirstDimOfArry(ByVal srcArry As Variant, i As Long) As Variant
    Dim tmpArry() As Variant

    '一次元目と二次元目を交換
    tmpArry = WorksheetFunction.Transpose(srcArry)

    ReDim Preserve tmpArry(UBound(tmpArry, 1), UBound(tmpArry, 2) + i)

    ExpandFirstDimOfArry = WorksheetFunction.Transpose(tmpArry)

End Function

'*************************************************************************************
'* Description  :二次元配列の指定行を削除する
'*************************************************************************************
Public Function DelRowFromArry(srcArry As Variant, delRowIdx As Long) As Variant
    
    Dim minRowIdx As Long, maxRowIdx As Long, minColIdx As Long, maxColIdx As Long
    Dim tmpArry As Variant
    Dim i As Long, j As Long, k As Long
    
    minRowIdx = LBound(srcArry, 1)
    maxRowIdx = UBound(srcArry, 1)
    minColIdx = LBound(srcArry, 2)
    maxColIdx = UBound(srcArry, 2)
    
    '元配列の行数-1で仮配列を再定義
    ReDim tmpArry(minRowIdx To maxRowIdx - 1, minColIdx To maxColIdx) As Variant
    
    k = minRowIdx
    For i = minRowIdx To maxRowIdx - 1
        '削除対象行であればk+1をして飛ばす
        If i = delRowIdx Then
            k = k + 1
        End If
        For j = minColIdx To maxColIdx
            tmpArry(i, j) = srcArry(k, j)
        Next j
        k = k + 1
    Next i
    
    DelRowFromArry = tmpArry
    
End Function

'********************************************************************************************
'* Description  :CSVファイル（タイトル行あり）のデータを二次元配列に格納して返す
'********************************************************************************************
Public Function LoadDataFromCSV(filePath As String) As Variant

    Const CHAR_SET = "UTF-8"
    Const NEWLINE_CHAR = vbCrLf
    Const SEPARATOR = ","

    Dim objADO As Object
    Dim buf As String, rowArry() As String, colArry() As String, titleArry() As String
    Dim rowLastIdx As Long
    Dim resArry() As Variant

    Dim i As Long, j As Long
    
    On Error GoTo ErrHandler:
    
    Set objADO = CreateObject("ADODB.Stream")
    
    objADO.Charset = CHAR_SET
    'objADO.Mode = 1  '読み取り専用
    objADO.Open
    '読み取り専用で開くと以下がエラーとなる（3219 このコンテキストで操作は許可されていません）
    objADO.LoadFromFile filePath
    
    'ファイルの一行目 (タイトル行）を読込み、区切り文字で区切ってtitleArryに格納
    buf = objADO.ReadText(-2)
    titleArry = Split(buf, SEPARATOR)
    
    rowLastIdx = 0
    ReDim rowArry(rowLastIdx) As String
    
    'ファイルの二行目から一行ずつ読込み、配列rowArryに格納
    Do Until objADO.EOS
        buf = objADO.ReadText(-2)
        ReDim Preserve rowArry(rowLastIdx)
        rowArry(rowLastIdx) = buf
        rowLastIdx = rowLastIdx + 1
    Loop
    
    ReDim resArry(UBound(rowArry), UBound(titleArry)) As Variant
    
    '配列rowArryの各要素を区切り文字で区切って、二次元配列resArryに格納
    For i = 0 To UBound(rowArry)
        'もし空行であれば、区切り文字を付加する
        If InStr(rowArry(i), SEPARATOR) = 0 Then
            rowArry(i) = SEPARATOR
        End If
        colArry = Split(rowArry(i), SEPARATOR)
        For j = 0 To UBound(titleArry)
            resArry(i, j) = colArry(j)
        Next j
    Next i

    LoadDataFromCSV = resArry
    
    GoTo Finally
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume Finally

Finally:
    objADO.Close
    
End Function

'********************************************************************************************
'* Description  :二次元配列のデータをCSVファイルへ出力して保存する
'********************************************************************************************
Public Function ExportArryToCSV(filePath As String, srcArry As Variant)

    Const CHAR_SET = "UTF-8"
    'Const CHAR_SET = "Shift-JIS"
    'Const NEWLINE_CHAR = vbCrLf
    Const SEPARATOR = ","
    
    'ADODB.Streamオブジェクトの定数定義（参照設定すれば定義の必要なし）
    Const adCRLF = -1                            '改行コードはCrLf
    Const adWriteLine = 1                        '行の終わりに指定した改行コードを書き込む
    Const adSaveCreateOverWrite = 2              '上書き保存OK

    Dim objADO As Object
    Dim i As Long, j As Long
    Dim buf As String

    Set objADO = CreateObject("ADODB.Stream")
    
    With objADO
        .Charset = CHAR_SET
        .LineSeparator = adCRLF
        .Open
        
        For i = LBound(srcArry, 1) To UBound(srcArry, 1)
            buf = ""
            For j = LBound(srcArry, 2) To UBound(srcArry, 2)
                '一行の文字列を作成
                If j <> UBound(srcArry, 2) Then
                    buf = buf & srcArry(i, j) & SEPARATOR
                Else
                    buf = buf & srcArry(i, j)
                End If
            Next j
            '一行書き出し
            .WriteText buf, adWriteLine
        Next i
        
        On Err GoTo ErrHandler
        'CSVファイルを上書き保存
        .SaveToFile filePath, adSaveCreateOverWrite
        
        .Close
        
    End With

    Exit Function
    
ErrHandler:
    Err.Raise Number:=601, Description:="ファイルの書込みに失敗しました。"

End Function

'********************************************************************************************
'* Description  :HTTP GET用にURLクエリパラメータをエンコードする
'********************************************************************************************
'Public Function encodeParams(pDic As Dictionary) As String
'
'    Dim pArry() As String
'    ReDim pArry(pDic.Count - 1) As String
'    Dim i As Long
'
'    For i = 0 To pDic.Count - 1
'        pArry(i) = pDic.Keys(i) & "=" & Application.WorksheetFunction.EncodeURL(pDic.Items(i))
'    Next i
'
'    encodeParams = Join(pArry, "&")
'
'End Function


'********************************************************************************************
'* Description  :URLからファイルをダウンロードする
'********************************************************************************************
Public Function DownloadFile(url As String, saveDir As String)
    
    Dim res As Long
    Dim fileName As String, filePath As String
    
    '指定されたダウンロード先ディレクトリが存在しなければ作成
    If Dir(saveDir, vbDirectory) = "" Then
        MkDir saveDir
    End If
    
    'URLからファイル名取得
    fileName = GetFileNameFromURL(url)
    
    filePath = saveDir & "\" & fileName
    
    'ダウンロード先ディレクトリに既に指定の名前のファイルがあればごみ箱へ移動
    If Dir(filePath) <> "" Then
        Call MoveFileToTrash(filePath)
    End If
        
    'ダウンロード実行
    res = URLDownloadToFile(0, url, filePath, 0, 0)
    
    If res = 0 Then
        Debug.Print "ダウンロード完了"
    Else
        Err.Raise Number:=602, Description:="ファイルをダウンロードできませんでした"
    End If
    
End Function

'********************************************************************************************
'* Description  :URLからファイル名を取り出す
'********************************************************************************************
Public Function GetFileNameFromURL(url As String) As String
    Dim tmp As Variant, tmpLastIdx As Long
    Dim fileName As String
    
    tmp = Split(url, "/")
    tmpLastIdx = UBound(tmp)
    fileName = tmp(tmpLastIdx)
    
    GetFileNameFromURL = fileName
    
End Function

'********************************************************************************************
'* Description  :URLからドメイン名を取り出す
'********************************************************************************************
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

'********************************************************************************************
'* Description  :ファイルをごみ箱へ移動する
'********************************************************************************************
Public Function MoveFileToTrash(filePath As String)
    
    Dim SH As SHFILEOPSTRUCT
    Dim res As Long
    
    With SH
        .hwnd = Application.hwnd
        .wFunc = FO_DELETE
        .pFrom = filePath
        .fFlags = FOF_ALLOWUNDO
    End With
    
    res = SHFileOperation(SH)
    
    If res <> 0 Then
        Err.Raise Number:=603, Description:="ごみ箱へ移動できませんでした"
    End If

End Function


