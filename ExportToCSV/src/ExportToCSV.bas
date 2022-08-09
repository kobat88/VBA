Attribute VB_Name = "Main"
'Option Explicit

Public Sub main()

    Const SRC_WS_NAME = "KO (3)"                 '�����ΏۃV�[�g��
    Const TITLE_ROW = 7                          '�^�C�g���s
    Const COL_LIMIT = 26                         '��̊g������iZ��j
    Const FIRST_COL = 1                          '�J�n��
    Const SAVE_DIR = "C:\Users\kobat88\Desktop\VBA\ExportToCSV" 'CSV��ۑ�����f�B���N�g��
    
    Dim lastCol As Long
    Dim objUtil As Util
    Dim resArry() As Variant
    Dim i As Long
    Dim objFSO As Object
    Dim csvFileName As String, saveFilePath As String
    
    Debug.Print "Main�J�n"
    
    On Err GoTo ErrHandler
    
    Set objUtil = New Util
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    '�ǂ̗�܂ō��ڂ����邩���ׂ�
    lastCol = ThisWorkbook.Worksheets(SRC_WS_NAME).Cells(TITLE_ROW, COL_LIMIT).End(xlToLeft).Column

    '�ΏۃZ���͈́i�^�C�g���s�������j�̃f�[�^��z��Ɋi�[
    resArry = objUtil.LoadDataFromSheet(ThisWorkbook, SRC_WS_NAME, TITLE_ROW + 1, FIRST_COL, lastCol)
        
        
    '�e���ڂ�����
    
    '�}�Ԃ��[�����߂�3���ɂ���
    For i = LBound(resArry, 1) To UBound(resArry, 1)
        resArry(i, 0) = Format(resArry(i, 0), "000")
    Next i
    
    '���i��������s���������A""�ň͂�
    For i = LBound(resArry, 1) To UBound(resArry, 1)
        resArry(i, 1) = Replace(resArry(i, 1), vbLf, "")
        resArry(i, 1) = """" & resArry(i, 1) & """"
    Next i
    
    
    '�z���CSV�t�@�C���֏o�͂��ĕۑ�
    csvFileName = objFSO.GetBaseName(ThisWorkbook.Name) & "_clean.csv"
    saveFilePath = SAVE_DIR & "\" & csvFileName
    Call objUtil.ExportArryToCSV(saveFilePath, resArry)
    
    Debug.Print "Main�I��"
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number, Err.Source, Err.Description

End Sub

