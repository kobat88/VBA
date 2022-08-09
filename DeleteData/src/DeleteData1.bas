Attribute VB_Name = "DeleteData1"
'Option Explicit

'Excel��Ń\�[�g -> �z����Ō������ăt���O�𗧂Ă� -> Excel�ɖ߂��ăt���O�Ńt�B���^���č폜
'�V�[�g�P�F1000�� �V�[�g�Q�F8000����3.5�b

Public Sub DeleteData1()

    Const WS_NAME1 = "Sheet1"
    Const WS_NAME2 = "Sheet2"
    Const WS_NAME1_SORT = "Sheet1_sort"
    Const WS_NAME2_SORT = "Sheet2_sort"

    Dim objUtil As Util
    Dim lastRow1 As Long, lastRow2 As Long
    Dim Arry1() As Variant, Arry2() As Variant
    Dim minRowIdx As Long, maxRowIdx As Long
    Dim i As Long, j As Long, csrJ As Long, k As Long
    Dim startTime As Single, endTime As Single, procTime As Single
    Dim startTime1 As Single, endTime1 As Single
    Dim startTime2 As Single, endTime2 As Single
    Dim startTime3 As Single, endTime3 As Single
    Dim startTime4 As Single, endTime4 As Single

    startTime = Timer
    
    Debug.Print "�����J�n"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set objUtil = New Util
    
    '�V�[�g�P�C�V�[�g�Q�̍ŏI�s���擾
    lastRow1 = Worksheets(WS_NAME1).Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = Worksheets(WS_NAME2).Cells(Rows.Count, 1).End(xlUp).Row
    
    
    startTime1 = Timer
    '�V�[�g�P�f�[�^���\�[�g�p�V�[�g�P�ɃR�s�[���ă\�[�g
    Worksheets(WS_NAME1).Copy After:=Worksheets(WS_NAME1)
    ActiveSheet.Name = WS_NAME1_SORT
    
    With Worksheets(WS_NAME1_SORT)
        '���̕��я���9��ڂɋL�^
        For k = 2 To lastRow1
            .Cells(k, 9).Value = k
        Next
        '1��ڂŃ\�[�g
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), Header:=xlYes
    End With
    
    '�V�[�g�Q�f�[�^���\�[�g�p�V�[�g�Q�ɃR�s�[���ă\�[�g
    Worksheets(WS_NAME2).Copy After:=Worksheets(WS_NAME2)
    ActiveSheet.Name = WS_NAME2_SORT
    
    With Worksheets(WS_NAME2_SORT)
        '���̕��я���9��ڂɋL�^
        For k = 2 To lastRow2
            .Cells(k, 9).Value = k
        Next
        '1��ڂŃ\�[�g
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), Header:=xlYes
    End With
    endTime1 = Timer
    Debug.Print "�\�[�g����:" & endTime1 - startTime1
    

    '�e�\�[�g�p�V�[�g�̃f�[�^��z��Ɋi�[
    startTime2 = Tiemr
    Arry1 = objUtil.LoadDataFromSheet(ThisWorkbook, WS_NAME1_SORT, 2, 1, 8)
    Arry2 = objUtil.LoadDataFromSheet(ThisWorkbook, WS_NAME2_SORT, 2, 1, 10)
    endTime2 = Timer
    Debug.Print "�z��֊i�[����:" & endTime2 - startTime2


    '�����ꂩ�̔z�񂪋�ł���Ώ����I��
    If objUtil.IsEmptyArry(Arry1) Or objUtil.IsEmptyArry(Arry2) Then
        Debug.Print "��r�Ώۃf�[�^����ł�"
        Exit Sub
    End If

    minRowIdx = LBound(Arry2, 1)
    maxRowIdx = UBound(Arry2, 1)

    '���C������
    startTime3 = Timer
    csrJ = LBound(Arry1, 1)
    '�z��Q�̐擪�s���珇�ɏ���
    For i = LBound(Arry2, 1) To UBound(Arry2, 1)
        '10��ڂ�"not exist"���i�[
        Arry2(i, 9) = "not exist"
        '�z��P�̐擪�s���珇�Ƀ`�F�b�N
        For j = csrJ To UBound(Arry1, 1)
            '�z��Q��A��=�z��P��A��ł���΁A�z��Q��10��ڂ�"exist"�ɍX�V���A�z��Q�̎��̍s��
            If Arry2(i, 0) = CStr(Arry1(j, 0)) Then
                Arry2(i, 9) = "exist"
                csrJ = j
                Exit For
            End If
        Next j
    Next i
    endTime3 = Timer
    Debug.Print "���C������:" & endTime3 - startTime3

    '�\�[�g�p�V�[�g�Q�̊����f�[�^�s���폜
    Worksheets(WS_NAME2_SORT).Rows("2:" & Cells.Rows.Count).Delete

    '�z��Q�̃f�[�^���\�[�g�p�V�[�g�Q�ɏo��
    startTime4 = Timer
    Call objUtil.OutDataToSheet(ThisWorkbook, WS_NAME2_SORT, 2, 1, Arry2)
    endTime4 = Timer
    Debug.Print "�V�[�g�ւ̏o�͏���:" & endTime4 - startTime4
    
    '�\�[�g�p�V�[�g�Q��10���="exist"�Ńt�B���^�[
    Worksheets(WS_NAME2_SORT).Range("A1").AutoFilter 10, "exist"
    
    '�t�B���^�[�����s���폜
    With Worksheets(WS_NAME2_SORT).Range("A1").CurrentRegion
        .Resize(.Rows.Count - 1).Offset(1, 0).Delete
    End With
    
    '�t�B���^�[����
    Worksheets(WS_NAME2_SORT).Range("A1").AutoFilter
    
    '�\�[�g�p�V�[�g�Q�̃f�[�^�����̕��я��Ń\�[�g
    Worksheets(WS_NAME2_SORT).Range("A1").CurrentRegion.Sort Key1:=Worksheets(WS_NAME2_SORT).Range("I1"), Header:=xlYes

    '�\�[�g�p�V�[�g�Q��9�10��ڂ��N���A
    Worksheets(WS_NAME2_SORT).Columns(9).Clear
    Worksheets(WS_NAME2_SORT).Columns(10).Clear

    '�V�[�g�Q���폜
    Worksheets(WS_NAME2).Delete

    '�\�[�g�p�V�[�g�Q���V�[�g�Q�փR�s�[
    Worksheets(WS_NAME2_SORT).Copy After:=Worksheets(WS_NAME1)
    ActiveSheet.Name = WS_NAME2

    '�\�[�g�p�V�[�g���폜
    Worksheets(WS_NAME1_SORT).Delete
    Worksheets(WS_NAME2_SORT).Delete
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = False

    Debug.Print "�����I��"
    
    endTime = Timer
    procTime = endTime - startTime
    
    Debug.Print "��������:" & procTime

End Sub

