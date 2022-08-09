Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Public Sub Main()

    Const SRC_WS_NAME = "�V�[�g�A"
    Const TGT_WS_NAME = "�V�[�g�@"

    Dim objUtil As Util_OptionBase1
    Dim srcArry() As Variant, tgtArry() As Variant
    Dim i As Long, j As Long, k As Long, updateCnt As Long
    Dim startTime As Single, endTime As Single
    
    Set objUtil = New Util_OptionBase1
    
    startTime = Timer

    Debug.Print "Main�J�n"

    On Err GoTo ErrHandler

    Application.ScreenUpdating = False

    '�R�s�[���V�[�g�̃f�[�^���R�s�[���z��Ɋi�[
    srcArry = objUtil.LoadDataFromSheet(ThisWorkbook, SRC_WS_NAME, 2, 1, 8)

    '�R�s�[��V�[�g�̊����f�[�^���R�s�[��z��Ɋi�[
    tgtArry = objUtil.LoadDataFromSheet(ThisWorkbook, TGT_WS_NAME, 2, 1, 8)
    
    
    '�R�s�[���z�񂪋�łȂ���Ώ�������
    If Not objUtil.IsEmptyArry(srcArry) Then
    
        '�R�s�[��z��̍ŏI�s���擾
        If objUtil.IsEmptyArry(tgtArry) Then
            k = 1
        Else
            k = UBound(tgtArry, 1) + 1
        End If
        
    
        '�R�s�[���z�����s������
        For i = LBound(srcArry, 1) To UBound(srcArry, 1)
            
            '�R�s�[���z��̈Č�ID�ŃR�s�[��z����������A�q�b�g������Y���s���X�V
            updateCnt = 0
            For j = LBound(tgtArry, 1) To UBound(tgtArry, 1)
                If tgtArry(j, 8) = srcArry(i, 1) Then
                    tgtArry(j, 1) = srcArry(i, 2)
                    tgtArry(j, 2) = srcArry(i, 3)
                    tgtArry(j, 3) = srcArry(i, 4)
                    tgtArry(j, 4) = srcArry(i, 5)
                    tgtArry(j, 5) = srcArry(i, 6)
                    tgtArry(j, 6) = srcArry(i, 7)
                    tgtArry(j, 7) = srcArry(i, 8)
                    updateCnt = updateCnt + 1
                End If
            Next j
            
            '�R�s�[���z��̈Č�ID���R�s�[��z��ɑ��݂��Ȃ���΁A�R�s�[��z��̍ŏI�s�̎��̍s�ɊY���s��ǉ�
            If updateCnt = 0 Then
            
                If objUtil.IsEmptyArry(tgtArry) Then
                    '�R�s�[��z�񂪋�Ȃ�΁A1�s�ōĒ�`
                    ReDim tgtArry(1, 8) As Variant
                Else
                    '�R�s�[��z�񂪋�łȂ���΁A�R�s�[��z��̍s�����P���₷
                    tgtArry = objUtil.ExpandFirstDimOfArry(tgtArry, 1)
                End If
            
                tgtArry(k, 1) = srcArry(i, 2)
                tgtArry(k, 2) = srcArry(i, 3)
                tgtArry(k, 3) = srcArry(i, 4)
                tgtArry(k, 4) = srcArry(i, 5)
                tgtArry(k, 5) = srcArry(i, 6)
                tgtArry(k, 6) = srcArry(i, 7)
                tgtArry(k, 7) = srcArry(i, 8)
                tgtArry(k, 8) = srcArry(i, 1)
                k = k + 1
            End If
        Next i
        
        '�R�s�[��z��̃f�[�^���V�[�g�@�֏o��
        Call objUtil.OutDataToSheet(ThisWorkbook, TGT_WS_NAME, 2, 1, tgtArry)
    
    Else
        Debug.Print SRC_WS_NAME & "�̃f�[�^��0���ł�"
    End If
    
    Set objUtil = Nothing
    
    Application.ScreenUpdating = True

    Debug.Print "Main�I��"
    
    endTime = Timer
    Debug.Print "Main:" & Now() & "," & (endTime - startTime) & "sec"
    
    Exit Sub

ErrHandler:
    Debug.Print Err.Number, Err.Description

End Sub

'*************************************************************************************
'* Description  :�񎟌����I�z��̃f�[�^��ێ������܂܈ꎟ���ڂ�v�f����i�������₷
'* �i�������z���ReDim Preserve�ł͍ŏI�����̗v�f���������₹�Ȃ����߁A���̊֐���p�Ӂj
'*************************************************************************************
'Public Function ExpandFirstDimOfArry(ByVal srcArry As Variant, i As Long) As Variant
'    Dim tmpArry() As Variant
'
'    '�ꎟ���ڂƓ񎟌��ڂ�����
'    tmpArry = WorksheetFunction.Transpose(srcArry)
'
'    ReDim Preserve tmpArry(UBound(tmpArry, 1), UBound(tmpArry, 2) + i)
'
'    ExpandFirstDimOfArry = WorksheetFunction.Transpose(tmpArry)
'
'End Function

