Attribute VB_Name = "GeoCoding"
Public Sub GeoCoding()

    Const WS_NAME = "Sheet1"

    Dim geoUrl As String
    Dim srcAddress As String
    Dim pDic As Dictionary
    Dim objReq As Object
    Dim jsonObj As Object

    Debug.Print "処理開始"

    '国土地理院API
    geoUrl = "https://msearch.gsi.go.jp/address-search/AddressSearch?"

    'JsonConverterを使用するためにはDictionaryの事前バインディング（参照設定）が必要
    'Set pDic = CreateObject("Scripting.Dictionary")
    Set pDic = New Dictionary

    'XMLHTTP.6.0としないと、objReq.readyStateがずっと1のまま（Sendが呼び出し済にならない）
    'Set objReq = CreateObject("MSXML2.XMLHTTP")
    Set objReq = CreateObject("MSXML2.XMLHTTP.6.0")

    'シートの緯度・経度をクリア
    With Worksheets(WS_NAME)
        .Cells(2, 2).Clear
        .Cells(2, 3).Clear

        '緯度経度を求める住所をシートから取得
        srcAddress = .Cells(2, 1).Value
    End With

    '住所をエンコードするためのディクショナリへ格納
    Call pDic.Add("q", srcAddress)

    '住所をエンコードしてパラメータにセットし、APIを実行
    '第三引数 True:非同期通信（省略値） False:同期通信　しかし、Trueを明示するとreadyStateが1のまま
    Call objReq.Open("GET", geoUrl & encodeParams(pDic))

    On Error GoTo ErrHandler
    'On Error Resume Next
    '送信不可の時、同期通信の場合は以下のsendでエラー、非同期通信の場合はsendでエラーにならない
    Call objReq.Send

    'On Error GoTo ErrHandler
    '同期通信の場合の送信エラー対応
    'If Err.Number <> 0 Then
    '    Err.Raise Number:=500, Description:="エラー：送信失敗"
    'End If

    waitStartTime = Timer
    'レスポンスが返るまで待つ
    Do While objReq.readyState < 4
        'タイムアウト設定
        If Timer - waitStartTime > 10 Then
            Err.Raise Number:=500, Description:="エラー：受信タイムアウト"
            Exit Do
        End If
        'DoEvents
        Debug.Print "readyState=" & objReq.readyState
    Loop

    On Error GoTo ErrHandler2
    'レスポンスJSONをパース
    Set jsonObj = JsonConverter.ParseJson(objReq.responseText)

    Debug.Print "緯度:" & jsonObj(1)("geometry")("coordinates")(2)
    Debug.Print "経度:" & jsonObj(1)("geometry")("coordinates")(1)

    '緯度・経度をシートに出力
    With Worksheets(WS_NAME)
        .Cells(2, 2) = jsonObj(1)("geometry")("coordinates")(2)
        .Cells(2, 3) = jsonObj(1)("geometry")("coordinates")(1)
    End With

Finally:
    Set pDic = Nothing
    Set objReq = Nothing
    
    Debug.Print "処理終了"
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number, Err.Description
    Resume Finally

ErrHandler2:
    Err.Description = "エラー：API失敗"
    Debug.Print Err.Number, Err.Description
    Resume Finally
    
End Sub

Public Function encodeParams(pDic As Dictionary) As String
    
    Dim pArry() As String
    ReDim pArry(pDic.Count - 1) As String
    Dim i As Long
    
    For i = 0 To pDic.Count - 1
        pArry(i) = pDic.Keys(i) & "=" & Application.WorksheetFunction.EncodeURL(pDic.Items(i))
    Next i
    
    encodeParams = Join(pArry, "&")

End Function

