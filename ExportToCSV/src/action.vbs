Dim objExcel,objWB

Set objExcel = WScript.CreateObject("Excel.Application")

'Excel��ʔ�\��
objExcel.Visible = False

'Excel�x����\��
objExcel.DisplayAlerts = False

'On Error Resume Next
set objWB = objExcel.Workbooks.Open(WScript.Arguments(0))

'If Err.Number <> 0 Then
'	WScript.Quit(1)
'End If

objExcel.Run WScript.Arguments(1)

'�ۑ����Ȃ��Ńt�@�C�������
objWB.Close False

Set objWB = Nothing
Set objExcel = Nothing
