Attribute VB_Name = "Main2"
'Option Explicit

'�������Ԕ�r����
'SRC_WS_MAME�̃f�[�^��=8000�s
'�z��iMain) => 0.2070313sec
'��d���[�v�iMain2�j=> 10.78516sec
'Find���\�b�h�iMain2�j => 10.64063sec
'Match���\�b�h�iMain2�j => 15.77344sec

Public Sub Main2()

    Const SRC_WS_NAME = "�V�[�g�A"
    Const TGT_WS_NAME = "�V�[�g�@"
    Const SRC_FIRST_ROW = 2
    Const TGT_FIRST_ROW = 2

    Dim srcFirstRow As Long, srcLastRow As Long, tgtFirstRow As Long, tgtLastRow As Long
    Dim i As Long, j As Long, updateCnt As Long
    Dim startTime As Single, endTime As Single
    Dim tgtRange As Object, tgtRow As Variant
    
    startTime = Timer
    
    Debug.Print "Main�J�n"
    
    Application.ScreenUpdating = False
    
    srcFirstRow = SRC_FIRST_ROW
    srcLastRow = Worksheets(SRC_WS_NAME).Cells(Rows.Count, 1).End(xlUp).Row
    tgtFirstRow = TGT_FIRST_ROW
    tgtLastRow = Worksheets(TGT_WS_NAME).Cells(Rows.Count, 1).End(xlUp).Row
    
    If srcFirstRow > srcLastRow Then
        Debug.Print SRC_WS_NAME & "�Ƀf�[�^������܂���"
        Exit Sub
    End If
    

    '��d���[�v����
    For i = srcFirstRow To srcLastRow
        updateCnt = 0
        For j = tgtFirstRow To tgtLastRow
            If Worksheets(TGT_WS_NAME).Cells(j, 8).Value = Worksheets(SRC_WS_NAME).Cells(i, 1).Value Then
                Worksheets(TGT_WS_NAME).Cells(j, 1).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
                Worksheets(TGT_WS_NAME).Cells(j, 2).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
                Worksheets(TGT_WS_NAME).Cells(j, 3).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
                Worksheets(TGT_WS_NAME).Cells(j, 4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
                Worksheets(TGT_WS_NAME).Cells(j, 5).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
                Worksheets(TGT_WS_NAME).Cells(j, 6).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
                Worksheets(TGT_WS_NAME).Cells(j, 7).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
                updateCnt = updateCnt + 1
            End If
        Next j
        If updateCnt = 0 Then
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 8).Value = Worksheets(SRC_WS_NAME).Cells(i, 1).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 1).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 2).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 3).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 5).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 6).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 7).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
            tgtLastRow = tgtLastRow + 1
        End If
    Next i


    'Find���\�b�h����
    '    For i = srcFirstRow To srcLastRow
    '        Set tgtRange = Worksheets(TGT_WS_NAME).Range(Cells(tgtFirstRow, 8), Cells(tgtLastRow, 8)).Find(Worksheets(SRC_WS_NAME).Cells(i, 1).Value)
    '        If Not tgtRange Is Nothing Then
    '            tgtRange.Offset(0, -1).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
    '            tgtRange.Offset(0, -2).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
    '            tgtRange.Offset(0, -3).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
    '            tgtRange.Offset(0, -4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
    '            tgtRange.Offset(0, -5).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
    '            tgtRange.Offset(0, -6).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
    '            tgtRange.Offset(0, -7).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
    '        Else
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 8).Value = Worksheets(SRC_WS_NAME).Cells(i, 1).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 1).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 2).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 3).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 5).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 6).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 7).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
    '            tgtLastRow = tgtLastRow + 1
    '        End If
    '    Next i
    
    
    'Match���\�b�h�����iWorksheetFunction.Math�ł͂Ȃ�Application.Match�j
    '    For i = srcFirstRow To srcLastRow
    '        Worksheets(TGT_WS_NAME).Activate
    '        If tgtLastRow < tgtFirstRow Then
    '            tgtRow = Application.Match(Worksheets(SRC_WS_NAME).Cells(i, 1).Value, Worksheets(TGT_WS_NAME).Range(Cells(tgtFirstRow, 8), Cells(tgtFirstRow, 8)), 0)
    '        Else
    '            tgtRow = Application.Match(Worksheets(SRC_WS_NAME).Cells(i, 1).Value, Worksheets(TGT_WS_NAME).Range(Cells(tgtFirstRow, 8), Cells(tgtLastRow, 8)), 0)
    '        End If
    '        If Not IsError(tgtRow) Then
    '            tgtRow = tgtFirstRow + tgtRow - 1
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 8).Value = Worksheets(SRC_WS_NAME).Cells(i, 1).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 1).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 2).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 3).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 5).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 6).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtRow, 7).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
    '        Else
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 8).Value = Worksheets(SRC_WS_NAME).Cells(i, 1).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 1).Value = Worksheets(SRC_WS_NAME).Cells(i, 2).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 2).Value = Worksheets(SRC_WS_NAME).Cells(i, 3).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 3).Value = Worksheets(SRC_WS_NAME).Cells(i, 4).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 4).Value = Worksheets(SRC_WS_NAME).Cells(i, 5).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 5).Value = Worksheets(SRC_WS_NAME).Cells(i, 6).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 6).Value = Worksheets(SRC_WS_NAME).Cells(i, 7).Value
    '            Worksheets(TGT_WS_NAME).Cells(tgtLastRow + 1, 7).Value = Worksheets(SRC_WS_NAME).Cells(i, 8).Value
    '            tgtLastRow = tgtLastRow + 1
    '        End If
    '    Next i

    Application.ScreenUpdating = True

    Debug.Print "Main�I��"
    
    endTime = Timer
    
    Debug.Print "Main2:" & Now() & "," & (endTime - startTime) & "sec"
    
    Exit Sub

End Sub

