Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()

    Const DL_FILE_URL = "https://www.mhlw.go.jp/content/pcr_positive_daily.csv"
    Const SAVE_DIR = "C:\Users\kobat88\Desktop\VBA\AnalyzeData\CovidDL"

    Dim objModel As Model
    Dim objUtil As Util
    Dim filePath As String
    Dim posCnts() As Variant
    
    Debug.Print "Main�J�n"
    
    On Error GoTo ErrHandler

    With Application
        .ScreenUpdating = False
    End With
    
    Set objModel = New Model
    Set objUtil = New Util
    
    '�����Ƃ̗z���Ґ���csv�t�@�C�����_�E�����[�h
    Call objUtil.DownloadFile(DL_FILE_URL, SAVE_DIR)
    
    filePath = SAVE_DIR & "\" & objUtil.GetFileNameFromURL(DL_FILE_URL)
    
    '�����Ƃ̗z���Ґ��̗݌v�����߂�Result�V�[�g�ɏo��
    Result.ProcName = "CalcPosTotal"
    Result.ProcResult = objModel.CalcPosTotal(filePath)
    Result.ProcDatetime = Now
    Result.ErrDesc = vbNullString
    Debug.Print "���ʏo�͊���"
    
    'Graph�V�[�g�ɃO���t�o��
    posCnts = objUtil.LoadDataFromCSV(filePath)
    Call Graph.DrawGraph(posCnts)
    
    Debug.Print "�O���t�o�͊���"
    
    GoTo Finally

ErrHandler:
    Result.ProcName = "CalcPosTotal"
    Result.ProcResult = 0
    Result.ProcDatetime = Now
    Result.ErrDesc = Err.Description

    Debug.Print Err.Number, Err.Source, Err.Description

    Resume Finally
    
Finally:
    With Application
        .ScreenUpdating = True
    End With
    
End Sub



Public Sub PrintSheet()

    Const PRT_WS_NAME = "Graph"

    With Worksheets(PRT_WS_NAME).PageSetup
        '�p��������
        .Orientation = xlLandscape
        '�]���ݒ�
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        '�Z���^�����O
        .CenterHorizontally = True
        .CenterVertically = True
        '�p���ꖇ�ɔ[�܂�悤�Ɉ��
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    'PageSetup�����������郁�\�b�h�������܂�ς��Ȃ�
    'Application.PrintCommunication = True
    
    Worksheets(PRT_WS_NAME).PrintOut Preview:=True

End Sub




