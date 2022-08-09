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


