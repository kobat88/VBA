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

'********************************************************************************************
'* Windows API�g�p�錾
'********************************************************************************************

'���ݔ��ֈړ�����API
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" _
(lpFileOp As SHFILEOPSTRUCT) As Long
'SHFileOperation�֐��ɓn�����[�U�[��`�^
Private Type SHFILEOPSTRUCT
    hwnd As Long                                 '�E�B���h�E�n���h��
    wFunc As Long                                '���s���鑀��
    pFrom As String                              '�Ώۃt�@�C����
    pTo As String                                '�ړI�t�@�C����
    fFlags As Integer                            '�t���O
    fAnyOperationsAborted As Long                '����
    hNameMappings As Long                        '�t�@�C�����}�b�s���O�I�u�W�F�N�g
    lpszProgressTitle As String                  '�_�C�A���O�̃^�C�g��
End Type

Private Const FO_DELETE = &H3                    '�폜����
Private Const FOF_ALLOWUNDO = &H40               '���ݔ��ɑ���

'URL����t�@�C���_�E�����[�h����API
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long


'****************************************************************************
'* �ȉ���Util��Function���`
'****************************************************************************

'****************************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂̃f�[�^��z��Ɋi�[���ĕԂ��iOption Base 0 �p�j
'* ��Function�́AOption Base 0 ���w�肵���i�ȗ��j���W���[���Ŏg�p����
'****************************************************************************************
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
            
            '�Z���͈͂��i�[�����z��͎d�l��Y������1�n�܂�ɂȂ�̂ŁA0�n�܂�̔z���ʓr��`
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

'*************************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂ɔz��̃f�[�^���o�͂���iOption Base 0 �p�j
'* ��Function�́AOption Base 0�i�ȗ��j ���w�肵�����W���[���Ŏg�p����
'*************************************************************************************
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

'****************************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂̃f�[�^��z��Ɋi�[���ĕԂ��iOption Base 1 �p�j
'* ��Function�́AOption Base 1 ���w�肵�����W���[���Ŏg�p����
'****************************************************************************************
Public Function LoadDataFromSheetBase1(srcWB As Workbook, srcWsName As String, firstRow As Long, firstCol As Long, lastCol As Long) As Variant

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

            '�w��̃Z���͈͂̃f�[�^��z��Ɋi�[
            resArry = Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
            
        End If
        
    End With
    
    LoadDataFromSheetBase1 = resArry
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'************************************************************************************
'* Description  :�w�肵���u�b�N�A�V�[�g�A�͈͂ɔz��̃f�[�^���o�͂���iOption Base 1 �p�j
'* ��Function�́AOption Base 1 ���w�肵�����W���[���Ŏg�p����
'************************************************************************************
Public Function OutDataToSheetBase1(tgtWB As Workbook, tgtWsName As String, firstRow As Long, firstCol As Long, srcArry() As Variant)

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
            
            lastRow = firstRow + UBound(srcArry, 1) - 1
            lastCol = firstCol + UBound(srcArry, 2) - 1
            
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

'*************************************************************************************
'* Description  :�񎟌����I�z��̃f�[�^��ێ������܂܈ꎟ���ڂ�v�f����i�������₷
'* �i�������z���ReDim Preserve�ł͍ŏI�����̗v�f���������₹�Ȃ����߁A���̊֐���p�Ӂj
'* Transpose�̎d�l��AOption Base 1�Ƃ���K�v����
'*************************************************************************************
Public Function ExpandFirstDimOfArry(ByVal srcArry As Variant, i As Long) As Variant
    Dim tmpArry() As Variant

    '�ꎟ���ڂƓ񎟌��ڂ�����
    tmpArry = WorksheetFunction.Transpose(srcArry)

    ReDim Preserve tmpArry(UBound(tmpArry, 1), UBound(tmpArry, 2) + i)

    ExpandFirstDimOfArry = WorksheetFunction.Transpose(tmpArry)

End Function

'*************************************************************************************
'* Description  :�񎟌��z��̎w��s���폜����
'*************************************************************************************
Public Function DelRowFromArry(srcArry As Variant, delRowIdx As Long) As Variant
    
    Dim minRowIdx As Long, maxRowIdx As Long, minColIdx As Long, maxColIdx As Long
    Dim tmpArry As Variant
    Dim i As Long, j As Long, k As Long
    
    minRowIdx = LBound(srcArry, 1)
    maxRowIdx = UBound(srcArry, 1)
    minColIdx = LBound(srcArry, 2)
    maxColIdx = UBound(srcArry, 2)
    
    '���z��̍s��-1�ŉ��z����Ē�`
    ReDim tmpArry(minRowIdx To maxRowIdx - 1, minColIdx To maxColIdx) As Variant
    
    k = minRowIdx
    For i = minRowIdx To maxRowIdx - 1
        '�폜�Ώۍs�ł����k+1�����Ĕ�΂�
        If i = delRowIdx Then
            k = k + 1
        End If
        For j = minColIdx To maxColIdx
            tmpArry(i, j) = srcArry(k, j)
        Next j
        k = k + 1
    Next i
    
    DelRowFromArry = tmpArry
    
End Function

'********************************************************************************************
'* Description  :CSV�t�@�C���i�^�C�g���s����j�̃f�[�^��񎟌��z��Ɋi�[���ĕԂ�
'********************************************************************************************
Public Function LoadDataFromCSV(filePath As String) As Variant

    Const CHAR_SET = "UTF-8"
    Const NEWLINE_CHAR = vbCrLf
    Const SEPARATOR = ","

    Dim objADO As Object
    Dim buf As String, rowArry() As String, colArry() As String, titleArry() As String
    Dim rowLastIdx As Long
    Dim resArry() As Variant

    Dim i As Long, j As Long
    
    On Error GoTo ErrHandler:
    
    Set objADO = CreateObject("ADODB.Stream")
    
    objADO.Charset = CHAR_SET
    'objADO.Mode = 1  '�ǂݎ���p
    objADO.Open
    '�ǂݎ���p�ŊJ���ƈȉ����G���[�ƂȂ�i3219 ���̃R���e�L�X�g�ő���͋�����Ă��܂���j
    objADO.LoadFromFile filePath
    
    '�t�@�C���̈�s�� (�^�C�g���s�j��Ǎ��݁A��؂蕶���ŋ�؂���titleArry�Ɋi�[
    buf = objADO.ReadText(-2)
    titleArry = Split(buf, SEPARATOR)
    
    rowLastIdx = 0
    ReDim rowArry(rowLastIdx) As String
    
    '�t�@�C���̓�s�ڂ����s���Ǎ��݁A�z��rowArry�Ɋi�[
    Do Until objADO.EOS
        buf = objADO.ReadText(-2)
        ReDim Preserve rowArry(rowLastIdx)
        rowArry(rowLastIdx) = buf
        rowLastIdx = rowLastIdx + 1
    Loop
    
    ReDim resArry(UBound(rowArry), UBound(titleArry)) As Variant
    
    '�z��rowArry�̊e�v�f����؂蕶���ŋ�؂��āA�񎟌��z��resArry�Ɋi�[
    For i = 0 To UBound(rowArry)
        '������s�ł���΁A��؂蕶����t������
        If InStr(rowArry(i), SEPARATOR) = 0 Then
            rowArry(i) = SEPARATOR
        End If
        colArry = Split(rowArry(i), SEPARATOR)
        For j = 0 To UBound(titleArry)
            resArry(i, j) = colArry(j)
        Next j
    Next i

    LoadDataFromCSV = resArry
    
    GoTo Finally
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume Finally

Finally:
    objADO.Close
    
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

'********************************************************************************************
'* Description  :HTTP GET�p��URL�N�G���p�����[�^���G���R�[�h����
'********************************************************************************************
'Public Function encodeParams(pDic As Dictionary) As String
'
'    Dim pArry() As String
'    ReDim pArry(pDic.Count - 1) As String
'    Dim i As Long
'
'    For i = 0 To pDic.Count - 1
'        pArry(i) = pDic.Keys(i) & "=" & Application.WorksheetFunction.EncodeURL(pDic.Items(i))
'    Next i
'
'    encodeParams = Join(pArry, "&")
'
'End Function


'********************************************************************************************
'* Description  :URL����t�@�C�����_�E�����[�h����
'********************************************************************************************
Public Function DownloadFile(url As String, saveDir As String)
    
    Dim res As Long
    Dim fileName As String, filePath As String
    
    '�w�肳�ꂽ�_�E�����[�h��f�B���N�g�������݂��Ȃ���΍쐬
    If Dir(saveDir, vbDirectory) = "" Then
        MkDir saveDir
    End If
    
    'URL����t�@�C�����擾
    fileName = GetFileNameFromURL(url)
    
    filePath = saveDir & "\" & fileName
    
    '�_�E�����[�h��f�B���N�g���Ɋ��Ɏw��̖��O�̃t�@�C��������΂��ݔ��ֈړ�
    If Dir(filePath) <> "" Then
        Call MoveFileToTrash(filePath)
    End If
        
    '�_�E�����[�h���s
    res = URLDownloadToFile(0, url, filePath, 0, 0)
    
    If res = 0 Then
        Debug.Print "�_�E�����[�h����"
    Else
        Err.Raise Number:=602, Description:="�t�@�C�����_�E�����[�h�ł��܂���ł���"
    End If
    
End Function

'********************************************************************************************
'* Description  :URL����t�@�C���������o��
'********************************************************************************************
Public Function GetFileNameFromURL(url As String) As String
    Dim tmp As Variant, tmpLastIdx As Long
    Dim fileName As String
    
    tmp = Split(url, "/")
    tmpLastIdx = UBound(tmp)
    fileName = tmp(tmpLastIdx)
    
    GetFileNameFromURL = fileName
    
End Function

'********************************************************************************************
'* Description  :�t�@�C�������ݔ��ֈړ�����
'********************************************************************************************
Public Function MoveFileToTrash(filePath As String)
    
    Dim SH As SHFILEOPSTRUCT
    Dim res As Long
    
    With SH
        .hwnd = Application.hwnd
        .wFunc = FO_DELETE
        .pFrom = filePath
        .fFlags = FOF_ALLOWUNDO
    End With
    
    res = SHFileOperation(SH)
    
    If res <> 0 Then
        Err.Raise Number:=603, Description:="���ݔ��ֈړ��ł��܂���ł���"
    End If

End Function

