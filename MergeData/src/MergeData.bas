Attribute VB_Name = "Main"
Option Explicit

'*************************************************************************************************
'* Description  :SRC_DIR���̑S�G�N�Z���t�@�C���̑S�V�[�g�̃f�[�^��TGT_SHEET_NAME�V�[�g�֏o�͂���
'*************************************************************************************************
Public Sub Main()

    Const SRC_DIR = "C:\Users\kobat88\Desktop\VBA\MergeData\TestDir\" '������"\"�܂ł���
    Const TGT_SHEET_NAME = "Merge"
    Const TITLE_ROW_NUM = 2
    Const FILE_NAME_COL_NUM = 2
    Const SHEET_NAME_COL_NUM = 3
    Const DATA_COL_NUM = 4
    
    Dim objUtil As Util
    Dim srcFileName As String
    Dim srcWB As Workbook
    Dim srcWS As Worksheet
    Dim resArry() As Variant
    Dim i As Long
    
    Debug.Print "Main�J�n"
    
    Application.ScreenUpdating = False
    
    With ThisWorkbook.Sheets(TGT_SHEET_NAME)
        
        '�^�C�g���s��艺�̍s���폜
        .Rows(TITLE_ROW_NUM & ":" & Cells.Rows.Count).Delete
        
        Set objUtil = New Util
        
        On Err GoTo ErrHandler
        
        'SRC_DIR���ɂ���G�N�Z���t�@�C���̃t�@�C�������擾
        srcFileName = Dir(SRC_DIR & "*.xls?")
        
        i = TITLE_ROW_NUM + 1
        
        '�R�s�[���t�@�C�����P�t�@�C�����J���ď���
        Do While srcFileName <> ""
            Set srcWB = Workbooks.Open(fileName:=SRC_DIR & srcFileName, ReadOnly:=True, UpdateLinks:=0)

            '�t�@�C�����o��
            .Cells(i, FILE_NAME_COL_NUM).Value = srcWB.Name
            
            '�R�s�[���t�@�C���̑S�V�[�g�����Ԃɏ���
            For Each srcWS In srcWB.Worksheets
            
                '�V�[�g���o��
                .Cells(i, SHEET_NAME_COL_NUM).Value = srcWS.Name
                
                '�ΏۃV�[�g�̃f�[�^���擾
                resArry = objUtil.LoadDataFromSheet(srcWB, srcWS.Name, 2, 2, 3)
                
                '�擾�f�[�^�o��
                Range(.Cells(i, DATA_COL_NUM), .Cells(i + UBound(resArry, 1), DATA_COL_NUM + UBound(resArry, 2))) = resArry
                
                i = i + UBound(resArry, 1)
            Next srcWS
            
            srcWB.Close
            
            '���̃t�@�C�������擾
            srcFileName = Dir()
        Loop
    
    End With
    
    Application.ScreenUpdating = True
    
    Debug.Print "Main�I��"
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

