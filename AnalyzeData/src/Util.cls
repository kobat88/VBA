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

