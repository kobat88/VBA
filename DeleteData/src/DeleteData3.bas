Attribute VB_Name = "DeleteData3"
'Option Explicit

'Excel��ŃV�[�g�P�����\�[�g -> Excel���match�Ō������ăt���O�𗧂Ă� -> �t���O�Ńt�B���^���č폜
'�V�[�g�P�F1000�� �V�[�g�Q�F8000����5�b

Public Sub DeleteData3()

    Const WS_NAME1 = "Sheet1"
    Const WS_NAME2 = "Sheet2"
    Const WS_NAME1_SORT = "Sheet1_sort"

    Dim objUtil As Util
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, k As Long
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
    
    
    '���C������
    startTime3 = Timer1
    '�V�[�g�Q��O���珇�ɏ���
    For i = 2 To lastRow2
        '9��ڂ�"not exist"���i�[
        Worksheets(WS_NAME2).Cells(i, 9) = "not exist"
        '�V�[�g�Q��A��̒l�ŁA�\�[�g�p�V�[�g�P��A���match�����i������Ȃ���΃G���[���Ԃ�j
        matchIdx = Application.Match(Worksheets(WS_NAME2).Cells(i, 1), Worksheets(WS_NAME1_SORT).Range("A:A"), 0)
        '�������ʂ�������΁A�V�[�g�Q��10��ڂ�"exist"�ōX�V
        If Not IsError(matchIdx) Then
            Worksheets(WS_NAME2).Cells(i, 9) = "exist"
        End If
    Next i
    endTime3 = Timer
    Debug.Print "���C������:" & endTime3 - startTime3

    
    '�V�[�g�Q��10���="exist"�Ńt�B���^�[
    Worksheets(WS_NAME2).Range("A1").AutoFilter 9, "exist"

    '�t�B���^�[�����s���폜
    With Worksheets(WS_NAME2).Range("A1").CurrentRegion
        .Resize(.Rows.Count - 1).Offset(1, 0).Delete
    End With

    '�t�B���^�[����
    Worksheets(WS_NAME2).Range("A1").AutoFilter

    '�V�[�g�Q��9��ڂ��N���A
    Worksheets(WS_NAME2).Columns(9).Clear

    '�\�[�g�p�V�[�g�P���폜
    Worksheets(WS_NAME1_SORT).Delete

    Application.ScreenUpdating = True
    Application.DisplayAlerts = False

    Debug.Print "�����I��"
    
    endTime = Timer
    procTime = endTime - startTime
    
    Debug.Print "��������:" & procTime & "@" & Now()

End Sub

