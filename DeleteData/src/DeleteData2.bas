Attribute VB_Name = "DeleteData2"
'Option Explicit

'Excel��Ń\�[�g -> Excel��Ń��[�v���������ăt���O�𗧂Ă� -> �t���O�Ńt�B���^���č폜
'�V�[�g�P�F1000�� �V�[�g�Q�F8000����13�b

Public Sub DeleteData2()

    Const WS_NAME1 = "Sheet1"
    Const WS_NAME2 = "Sheet2"
    Const WS_NAME1_SORT = "Sheet1_sort"
    Const WS_NAME2_SORT = "Sheet2_sort"

    Dim objUtil As Util
    Dim lastRow1 As Long, lastRow2 As Long
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
    
    
    '���C������
    startTime3 = Timer1
    csrJ = 2
    '�\�[�g�p�V�[�g�Q��O���珇�ɏ���
    For i = 2 To lastRow2
        '10��ڂ�"not exist"���i�[
        Worksheets(WS_NAME2_SORT).Cells(i, 10) = "not exist"
        '�\�[�g�p�V�[�g�P��O���珇�Ƀ`�F�b�N
        For j = csrJ To lastRow1
            '�\�[�g�p�V�[�g�Q��A��=�\�[�g�p�V�[�g�P��A��ł���΁A�\�[�g�p�V�[�g�Q��10��ڂ�"exist"�ɍX�V���A�\�[�g�p�V�[�g�Q�̎��̍s��
            If Worksheets(WS_NAME2_SORT).Cells(i, 1) = Worksheets(WS_NAME1_SORT).Cells(j, 1) Then
                Worksheets(WS_NAME2_SORT).Cells(i, 10) = "exist"
                csrJ = j
                Exit For
            End If
        Next j
    Next i
    endTime3 = Timer
    Debug.Print "���C������:" & endTime3 - startTime3

    
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

