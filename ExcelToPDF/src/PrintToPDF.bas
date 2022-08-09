Attribute VB_Name = "PrintToPDF"
Public Sub PrintToPDF()

    Const WS_NAME = "Sheet1"

    Dim pdfFilePath As String

    pdfFilePath = ThisWorkbook.Path & "\test1.pdf"

    With ThisWorkbook.Worksheets(WS_NAME)

        '���y�[�W�ݒ�
        .Rows(2).PageBreak = xlPageBreakManual
    
        '���ɂ�����͈͂̐ݒ�ŗ]���Ȃǒ����ł���
    
        .PrintOut _
        ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, _
        PrToFileName:=pdfFilePath, _
        Preview:=False

    End With
        
End Sub


