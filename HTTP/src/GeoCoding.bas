Attribute VB_Name = "GeoCoding"
Public Sub GeoCoding()

    Const WS_NAME = "Sheet1"

    Dim geoUrl As String
    Dim srcAddress As String
    Dim pDic As Dictionary
    Dim objReq As Object
    Dim jsonObj As Object

    Debug.Print "�����J�n"

    '���y�n���@API
    geoUrl = "https://msearch.gsi.go.jp/address-search/AddressSearch?"

    'JsonConverter���g�p���邽�߂ɂ�Dictionary�̎��O�o�C���f�B���O�i�Q�Ɛݒ�j���K�v
    'Set pDic = CreateObject("Scripting.Dictionary")
    Set pDic = New Dictionary

    'XMLHTTP.6.0�Ƃ��Ȃ��ƁAobjReq.readyState��������1�̂܂܁iSend���Ăяo���ςɂȂ�Ȃ��j
    'Set objReq = CreateObject("MSXML2.XMLHTTP")
    Set objReq = CreateObject("MSXML2.XMLHTTP.6.0")

    '�V�[�g�̈ܓx�E�o�x���N���A
    With Worksheets(WS_NAME)
        .Cells(2, 2).Clear
        .Cells(2, 3).Clear

        '�ܓx�o�x�����߂�Z�����V�[�g����擾
        srcAddress = .Cells(2, 1).Value
    End With

    '�Z�����G���R�[�h���邽�߂̃f�B�N�V���i���֊i�[
    Call pDic.Add("q", srcAddress)

    '�Z�����G���R�[�h���ăp�����[�^�ɃZ�b�g���AAPI�����s
    '��O���� True:�񓯊��ʐM�i�ȗ��l�j False:�����ʐM�@�������ATrue�𖾎������readyState��1�̂܂�
    Call objReq.Open("GET", geoUrl & encodeParams(pDic))

    On Error GoTo ErrHandler
    'On Error Resume Next
    '���M�s�̎��A�����ʐM�̏ꍇ�͈ȉ���send�ŃG���[�A�񓯊��ʐM�̏ꍇ��send�ŃG���[�ɂȂ�Ȃ�
    Call objReq.Send

    'On Error GoTo ErrHandler
    '�����ʐM�̏ꍇ�̑��M�G���[�Ή�
    'If Err.Number <> 0 Then
    '    Err.Raise Number:=500, Description:="�G���[�F���M���s"
    'End If

    waitStartTime = Timer
    '���X�|���X���Ԃ�܂ő҂�
    Do While objReq.readyState < 4
        '�^�C���A�E�g�ݒ�
        If Timer - waitStartTime > 10 Then
            Err.Raise Number:=500, Description:="�G���[�F��M�^�C���A�E�g"
            Exit Do
        End If
        'DoEvents
        Debug.Print "readyState=" & objReq.readyState
    Loop

    On Error GoTo ErrHandler2
    '���X�|���XJSON���p�[�X
    Set jsonObj = JsonConverter.ParseJson(objReq.responseText)

    Debug.Print "�ܓx:" & jsonObj(1)("geometry")("coordinates")(2)
    Debug.Print "�o�x:" & jsonObj(1)("geometry")("coordinates")(1)

    '�ܓx�E�o�x���V�[�g�ɏo��
    With Worksheets(WS_NAME)
        .Cells(2, 2) = jsonObj(1)("geometry")("coordinates")(2)
        .Cells(2, 3) = jsonObj(1)("geometry")("coordinates")(1)
    End With

Finally:
    Set pDic = Nothing
    Set objReq = Nothing
    
    Debug.Print "�����I��"
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number, Err.Description
    Resume Finally

ErrHandler2:
    Err.Description = "�G���[�FAPI���s"
    Debug.Print Err.Number, Err.Description
    Resume Finally
    
End Sub

Public Function encodeParams(pDic As Dictionary) As String
    
    Dim pArry() As String
    ReDim pArry(pDic.Count - 1) As String
    Dim i As Long
    
    For i = 0 To pDic.Count - 1
        pArry(i) = pDic.Keys(i) & "=" & Application.WorksheetFunction.EncodeURL(pDic.Items(i))
    Next i
    
    encodeParams = Join(pArry, "&")

End Function

