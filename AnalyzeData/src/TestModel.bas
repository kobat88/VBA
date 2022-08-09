Attribute VB_Name = "TestModel"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    'Const FILE_PATH = "C:\Users\kobat88\Desktop\VBA\AnalyzeData\CovidDL\pcr_positive_daily.csv"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestCalcPosTotal()                   'TODO Rename test
    
    'TestInitialize()�Œ�`�����������܂������Ȃ�
    Const FILE_PATH = "C:\Users\kobat88\Desktop\VBA\AnalyzeData\CovidDL\pcr_positive_daily.csv"
    
    '�O���[�o���ϐ��ɂ���TestInitialize()��New�����������܂������Ȃ�
    Dim objModel As Model
    Dim posTotal As Long
    
    On Error GoTo TestFail
    
    Set objModel = New Model
    posTotal = objModel.CalcPosTotal(FILE_PATH)
    
    '���Ғl��Long�^�Ő錾���Ȃ��ƃe�X�g���ʂŌ^�Ⴂ�̌x���ƂȂ�
    Dim expValue As Long
    expValue = 875231
    
    'Arrange:

    'Act:

    'Assert:
    'Assert.Succeed
    Assert.AreEqual expValue, posTotal

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

