Attribute VB_Name = "CombineData"
Public Sub Main()

    Dim srcArry1() As Variant, srcArry2() As Variant, resArry() As Variant
    Dim i As Long, j As Long, k As Long, combiLen As Long
    Dim objUtil As Util
    
    Debug.Print "Main�J�n"
    
    Set objUtil = New Util
    
    '�g�ݍ��킹��f�[�^��񂲂Ƃɔz��Ɋi�[
    srcArry1 = objUtil.LoadDataFromSheet(ThisWorkbook, "Sheet1", 2, 1, 1)
    srcArry2 = objUtil.LoadDataFromSheet(ThisWorkbook, "Sheet1", 2, 2, 2)
    
    '�g�ݍ��킹��f�[�^�̂����ꂩ�̗񂪋�̏ꍇ�G���[
    If objUtil.IsEmptyArry(srcArry1) Or objUtil.IsEmptyArry(srcArry2) Then
        Debug.Print "�G���[�F�f�[�^��̂����ꂩ����ł�"
        Exit Sub
    End If
    
    '�g�ݍ��킹�̐����擾
    combiLen = (UBound(srcArry1) + 1) * (UBound(srcArry2) + 1)
    
    '�s�����g�ݍ��킹�̐��Ō��ʂ��i�[����z����Ē�`
    ReDim resArry(combiLen - 1, 0) As Variant

    '�g�ݍ��킹�𐶐����Č��ʔz��Ɋi�[
    k = 0
    For i = LBound(srcArry1, 1) To UBound(srcArry1, 1)
        For j = LBound(srcArry2, 1) To UBound(srcArry2, 1)
            resArry(k, 0) = srcArry2(j, 0) & "��" & srcArry1(i, 0)
            k = k + 1
        Next j
    Next i

    '���ʔz����V�[�g�ɏo��
    Call objUtil.OutDataToSheet(ThisWorkbook, "sheet1", 2, 3, resArry)

    Debug.Print "Main�I��"

End Sub

