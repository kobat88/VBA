Dim objExcel,objWB

Set objExcel = WScript.CreateObject("Excel.Application")

'Excel画面非表示
objExcel.Visible = False

'Excel警告非表示
objExcel.DisplayAlerts = False

'On Error Resume Next
set objWB = objExcel.Workbooks.Open(WScript.Arguments(0))

'If Err.Number <> 0 Then
'	WScript.Quit(1)
'End If

objExcel.Run WScript.Arguments(1)

'保存しないでファイルを閉じる
objWB.Close False

Set objWB = Nothing
Set objExcel = Nothing
