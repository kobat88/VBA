Attribute VB_Name = "Main"
'Option Explicit

Public Sub Main()

    Const SRC_WS_NAME = "Sheet1"
    Const PPT_FILE_PATH = "C:\Users\kobat88\Desktop\VBA\ExcelToPPT\ExcelToPPT.pptx"
    Const LAYOUT_NAME = "����Ɖ�"

    Dim objPPT As Object
    Dim pptFile As Object
    Dim objOffice As Object
    Dim pptFilePath As Variant, fileFilterStr As String, titleStr As String
    Dim objFSO As Object
    Dim objPres As Object
    Dim presSlide As Object
    Dim cl As Object
    Dim clIdx As Long
    Dim objLayout As Object
    Dim objSlide As Object
    Dim shp As Object
    Dim i As Long

    Debug.Print "Main�J�n"

    Set objPPT = CreateObject("PowerPoint.Application")
    Set objFSO = CreateObject("scripting.FileSystemObject")

    On Error GoTo ErrHandler1

    '�V�K�v���[���e�[�V�����t�@�C���쐬
    'Set objPres = objPPT.Presentations.Add()

    '�w��̃t�@�C�������݂��Ȃ���΃G���[
    If Dir(PPT_FILE_PATH) = "" Then
        Err.Raise Number:=500, Description:="�w��̃t�@�C�������݂��܂���"
    End If

    '�w��̃t�@�C�������ɊJ����Ă�����G���[
    For Each pptFile In objPPT.Presentations
        If pptFile.FullName = PPT_FILE_PATH Then
            Err.Raise Number:=500, Description:="�w��̃t�@�C�����J����Ă��܂�"
        End If
    Next

    '�t�@�C�����J���_�C�A���O�\��
    '�J�����g�t�H���_�ȊO�������\���������ꍇ�́AChDir �t�H���_ ����
    fileFilterStr = "PowerPoint �v���[���e�[�V����,*.pptx"
    titleStr = "���f��̃p���[�|�C���g�t�@�C����I�����Ă�������"
    'pptFilePath = Application.GetOpenFilename(FileFilter = fileFilterStr, Title = titleStr)
    pptFilePath = Application.GetOpenFilename(fileFilterStr, , titleStr)
    If pptFilePath = False Then
        Err.Raise Number:=500, Description:="���f��̃t�@�C�����I������Ă��܂���"
    End If
    
    '�w��̃t�@�C�����o�b�N�A�b�v
    Call objFSO.CopyFile(pptFilePath, pptFilePath & Format(Now(), "yyyymmdd-hhmmss") & ".backup")
    
    '�w��̃t�@�C������ʔ�\���ŊJ��
    Set objPres = objPPT.Presentations.Open(pptFilePath, WithWindow:=msoFalse)

    On Error GoTo ErrHandler2

    '�����̃X���C�h��S�č폜
    For i = objPres.Slides.Count To 1 Step -1
        objPres.Slides(i).Delete
    Next

    '�X���C�h���C�A�E�g�̑I��
    clIdx = 0
    For Each cl In objPres.SlideMaster.CustomLayouts
        If cl.Name = LAYOUT_NAME Then
            clIdx = cl.Index
        End If
    Next
    If clIdx = 0 Then
        Err.Raise Number:=500, Description:="�w��̃��C�A�E�g������܂���"
    End If
    
    Set objLayout = objPres.SlideMaster.CustomLayouts(clIdx)

    '�V�F�C�v�ԍ��𒲂ׂ邽�߂̃f�o�b�O���
    'objPres.Slides.AddSlide 1, objLayout
    'Set objSlide = objPres.Slides(1)
    'i = 1
    'For Each shp In objSlide.Shapes
    '    Debug.Print shp.Name & ",IdxNo=" & i
    '    i = i + 1
    'Next shp
    '��
    'Title 1, IdxNo = 1
    'Text Placeholder 2,IdxNo=2
    'Text Placeholder 3,IdxNo=3

    On Error GoTo ErrHandler3
    '�G�N�Z���̃f�[�^���V�F�C�v�ɏ�������
    With ThisWorkbook.Worksheets(SRC_WS_NAME)
        For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row - 1
            objPres.Slides.AddSlide(i, objLayout).Shapes(2).TextFrame.TextRange.Text = .Cells(i + 1, 1).Value
            objPres.Slides(i).Shapes(3).TextFrame.TextRange.Text = .Cells(i + 1, 2).Value
        Next i
    End With

    On Error GoTo ErrHandler2
    '�t�@�C����ۑ�
    objPres.Save

    '�ۑ�����������܂ő҂�
    Do Until objPres.Saved
        Debug.Print "�v���[���e�[�V�����ۑ���"
    Loop

Finally:
    On Error Resume Next
    
    '�t�@�C�������
    objPres.Close
    
    '�p���[�|�C���g���I��
    objPPT.Quit
    Set objPPT = Nothing
    Set objFSO = Nothing
       
    Debug.Print "Main�I��"
        
    Exit Sub

ErrHandler1:
    Debug.Print Err.Number, Err.Description
    objPPT.Quit
    Set objPPT = Nothing
    Exit Sub

ErrHandler2:
    Debug.Print Err.Number, Err.Description
    Resume Finally
    
ErrHandler3:
    Err.Raise Number:=500, Description:="�V�F�C�v�ւ̏����݂Ɏ��s���܂���"
    Debug.Print Err.Number, Err.Description
    Resume Finally
   
End Sub

