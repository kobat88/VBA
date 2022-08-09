Attribute VB_Name = "TestUtil"
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
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestLoadDataFromSheetReturnEmpty()   'TODO Rename test
    On Error GoTo TestFail

    Dim objUtil As Util
    Dim resArry() As Variant

    Set objUtil = New Util

    resArry = objUtil.LoadDataFromSheet(ThisWorkbook, "test2", 100, 2, 6)

    'Arrange:

    'Act:

    'Assert:
    'Assert.Succeed
    
    Assert.IsTrue objUtil.IsEmptyArry(resArry)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestLoadDataFromSheet()              'TODO Rename test
    On Error GoTo TestFail
    
    Dim objUtil As Util
    Dim resArry() As Variant
    
    Set objUtil = New Util
    
    resArry = objUtil.LoadDataFromSheet(ThisWorkbook, "test2", 2, 1, 5)
    Call objUtil.OutDataToSheet(ThisWorkbook, "test1", 2, 1, resArry)
    
    'Arrange:

    'Act:

    'Assert:
    'Assert.Succeed
    
    'resArryの行数確認
    Assert.AreEqual CLng(5), UBound(resArry, 1) - LBound(resArry, 1) + 1
    
    'resArryの列数確認
    Assert.AreEqual CLng(5), UBound(resArry, 2) - LBound(resArry, 2) + 1
    
    'resArryの値の確認
    Assert.AreEqual "あ", ThisWorkbook.Worksheets("test1").Cells(2, 1).Value
    Assert.AreEqual #8/2/2022#, ThisWorkbook.Worksheets("test1").Cells(3, 3).Value
    Assert.AreEqual "\\", ThisWorkbook.Worksheets("test1").Cells(4, 5).Value
    Assert.AreEqual Nothing, ThisWorkbook.Worksheets("test1").Cells(6, 5).Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

''@TestMethod("Uncategorized")
'Private Sub TestLoadDataFromSheetErr()                        'TODO Rename test
'    Const ExpectedError As Long = 501              'TODO Change to expected error number
'    On Error GoTo TestFail
'
'    'Arrange:
'
'    'Act:
'
'    Dim objUtil As Util
'    Dim resArry() As Variant
'
'    Set objUtil = New Util
'
'    resArry = objUtil.LoadDataFromSheet(ThisWorkbook, "test2", 0, 2, 6)
'
'Assert:
'    Assert.Fail "Expected error was not raised"
'
'TestExit:
'    Exit Sub
'TestFail:
'    If Err.Number = ExpectedError Then
'        Resume TestExit
'    Else
'        Resume Assert
'    End If
'End Sub


'@TestMethod("Uncategorized")
Private Sub TestOutDataToSheet()                 'TODO Rename test
    On Error GoTo TestFail
    
    Dim objUtil As Util
    Dim resArry() As Variant
    
    Set objUtil = New Util
    
    ReDim resArry(1, 1) As Variant
    resArry(0, 0) = "あ"
    resArry(1, 0) = "い"
    resArry(0, 1) = "a"
    resArry(1, 1) = "b"
    
    Call objUtil.OutDataToSheet(ThisWorkbook, "test3", 2, 1, resArry)
    
    'Arrange:

    'Act:

    'Assert:
    'Assert.Succeed
    
    '出力の確認
    Assert.AreEqual "あ", ThisWorkbook.Worksheets("test3").Cells(2, 1).Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestOutDataToSheetFromEmpty()        'TODO Rename test
    On Error GoTo TestFail
    
    Dim objUtil As Util
    Dim resArry() As Variant
    
    Set objUtil = New Util
    
    resArry = VBA.Array()
    
    Call objUtil.OutDataToSheet(ThisWorkbook, "test3", 2, 1, resArry)
    
    'Arrange:

    'Act:

    'Assert:
    'Assert.Succeed
    
    '出力の確認
    Assert.AreEqual Nothing, ThisWorkbook.Worksheets("test4").Cells(2, 1).Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

