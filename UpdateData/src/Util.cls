VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂̃f�[�^��z��Ɋi�[���ĕԂ�
'****************************************************************************
Public Function LoadDataFromSheet(srcWB As Workbook, srcWsName As String, firstRow As Long, firstCol As Long, lastCol As Long) As Variant

    Dim lastRow As Long
    Dim preArry() As Variant, resArry() As Variant
    Dim i As Long, j As Long
    
    On Err GoTo ErrHandler
    
    If firstRow <= 0 Or firstCol <= 0 Or firstCol > lastCol Then
        Err.Raise Number:=501, Description:="�J�n�s�A�J�n��A�I����̎w�肪�s���ł�"
    End If
    
    
    With srcWB.Worksheets(srcWsName)
    
        '�t�B���^���ݒ肳��Ă���Ή����i�I�[�g�t�B���^�̐ݒ�͎c���j
        If .FilterMode Then
            .ShowAllData
        End If
        
        '�ŏI�s�擾
        lastRow = .Cells(Rows.Count, firstCol).End(xlUp).Row
        
        
        If lastRow < firstRow Then
        
            '�ŏI�s���J�n�s��菬������Ζ߂�l�ɋ�̔z���ݒ�
            resArry = VBA.Array()
        
        Else
        
            '�w��̃Z���͈͂̃f�[�^���v���z��Ɋi�[
            preArry = Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
            
            '�Z���͈͂��i�[�����z��͉��̂��Y������1�n�܂�ɂȂ�̂ŁA0�n�܂�̔z���ʓr��`
            ReDim resArry(0 To lastRow - firstRow, 0 To lastCol - firstCol)
            
            '�v���z��̃f�[�^��ʓr��`�����z��Ɋi�[
            For i = 1 To UBound(preArry, 1)
                For j = 1 To UBound(preArry, 2)
                    resArry(i - 1, j - 1) = preArry(i, j)
                Next j
            Next i
            
        End If
        
    End With
    
    LoadDataFromSheet = resArry
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'****************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂ɔz��̃f�[�^���o�͂���
'****************************************************************************
Public Function OutDataToSheet(tgtWB As Workbook, tgtWsName As String, firstRow As Long, firstCol As Long, srcArry() As Variant)

    Dim lastRow As Long, lastCol As Long
    
    On Err GoTo ErrHandler
    
    If firstRow <= 0 Or firstCol <= 0 Then
        Err.Raise Number:=501, Description:="�J�n�s�A�J�n��̎w�肪�s���ł�"
    End If
    
    '�n���ꂽ�z�񂪋�łȂ���Ώ�������
    If Not IsEmptyArry(srcArry) Then
    
        With tgtWB.Worksheets(tgtWsName)
        
            '�t�B���^���ݒ肳��Ă���Ή����i�I�[�g�t�B���^�̐ݒ�͎c���j
            If .FilterMode Then
                .ShowAllData
            End If
            
            lastRow = firstRow + UBound(srcArry, 1)
            lastCol = firstCol + UBound(srcArry, 2)
            
            Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol)) = srcArry
            
        End With
        
    End If
        
    Exit Function
        
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'****************************************************************************
'* Description  :�z�񂪋󂩂ǂ����𔻒肷��i��̏ꍇTrue��Ԃ��j
'****************************************************************************
Public Function IsEmptyArry(myArry() As Variant) As Boolean

    If 0 <= UBound(myArry, 1) Then
        IsEmptyArry = False
    Else
        IsEmptyArry = True
    End If
    
End Function
