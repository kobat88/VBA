Attribute VB_Name = "DistinctData"
'�f�[�^64,000���ŏ�������13�b

Public Sub DistinctData()

Const WS_NAME = "Sheet1"

Dim lastRow As Long
Dim i As Long

startTime = Timer
Debug.Print "�����J�n"

Application.ScreenUpdating = False

With Worksheets(WS_NAME)

    '�ŏI�s���擾
    lastRow = .Cells(Rows.Count, 1).End(xlUp).Row

    '���̕��я���2��ڂɋL�^
    For i = 1 To lastRow
        .Cells(i, 2).Value = i
    Next i

    '1��ځiURL�j�Ń\�[�g
    .Range("A1").CurrentRegion.Sort Key1:=.Range("A1")

    '�ォ�珇�Ɏ��̍s�Ɣ�ׂăh���C���������Ȃ玟�̍s�𒅐F
    For i = 1 To lastRow
        If ExtractDomain(.Cells(i, 1).Value) = ExtractDomain(.Cells(i + 1, 1).Value) Then
            .Cells(i + 1, 1).Interior.Color = RGB(200, 200, 200)
        End If
    Next i
    
    '���̏��ōă\�[�g
    .Range("A1").CurrentRegion.Sort Key1:=.Range("B1")
    
    '2��ڂ̒l���N���A
    .Columns(2).Clear
            
End With

Application.ScreenUpdating = True

Debug.Print "�����I��"
endTime = Timer
Debug.Print "�������ԁF" & endTime - startTime & "@" & Now()

End Sub


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
