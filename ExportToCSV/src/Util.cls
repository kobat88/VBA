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

'********************************************************************************************
'* Description  :�񎟌��z��̃f�[�^��CSV�t�@�C���֏o�͂��ĕۑ�����
'********************************************************************************************
Public Function ExportArryToCSV(filePath As String, srcArry As Variant)

    Const CHAR_SET = "UTF-8"
    'Const CHAR_SET = "Shift-JIS"
    'Const NEWLINE_CHAR = vbCrLf
    Const SEPARATOR = ","
    
    'ADODB.Stream�I�u�W�F�N�g�̒萔��`�i�Q�Ɛݒ肷��Β�`�̕K�v�Ȃ��j
    Const adCRLF = -1                            '���s�R�[�h��CrLf
    Const adWriteLine = 1                        '�s�̏I���Ɏw�肵�����s�R�[�h����������
    Const adSaveCreateOverWrite = 2              '�㏑���ۑ�OK

    Dim objADO As Object
    Dim i As Long, j As Long
    Dim buf As String

    Set objADO = CreateObject("ADODB.Stream")
    
    With objADO
        .Charset = CHAR_SET
        .LineSeparator = adCRLF
        .Open
        
        For i = LBound(srcArry, 1) To UBound(srcArry, 1)
            buf = ""
            For j = LBound(srcArry, 2) To UBound(srcArry, 2)
                '��s�̕�������쐬
                If j <> UBound(srcArry, 2) Then
                    buf = buf & srcArry(i, j) & SEPARATOR
                Else
                    buf = buf & srcArry(i, j)
                End If
            Next j
            '��s�����o��
            .WriteText buf, adWriteLine
        Next i
        
        On Err GoTo ErrHandler
        'CSV�t�@�C�����㏑���ۑ�
        .SaveToFile filePath, adSaveCreateOverWrite
        
        .Close
        
    End With

    Exit Function
    
ErrHandler:
    Err.Raise Number:=601, Description:="�t�@�C���̏����݂Ɏ��s���܂����B"

End Function
