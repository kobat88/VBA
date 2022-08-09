VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************************
'* Description  :日ごとのPCR検査陽性者数の累計を返す
'********************************************************************************************
Public Function CalcPosTotal(filePath As String) As Long

    Dim objUtil As Util
    Dim posCnts() As Variant
    Dim posTotal As Long
    Dim lngPosCnt As Long, i As Long
    
    Set objUtil = New Util
    
    '日ごとの陽性者数を配列で取得
    posCnts = objUtil.LoadDataFromCSV(filePath)
    
    '日ごとの陽性者数の累計を計算
    On Err GoTo ErrHandler
    posTotal = 0
    lngPosCnt = 0
    For i = 0 To UBound(posCnts)
        '1項目目が日付でかつ2項目目が数値であればLong型に変換して累計に計上
        If IsDate(posCnts(i, 0)) <> False And IsNumeric(posCnts(i, 1)) Then
            lngPosCnt = CLng(posCnts(i, 1))
            posTotal = posTotal + lngPosCnt
        Else
            Debug.Print (i + 1) & "行目のデータ:" & posCnts(i, 0) & "," & posCnts(i, 1)
            Err.Raise Number:=500, Description:=(i + 1) & "行目のデータが不正です"
        End If
    Next i
    
    CalcPosTotal = posTotal
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

