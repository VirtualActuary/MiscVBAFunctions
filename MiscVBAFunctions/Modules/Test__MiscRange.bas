Attribute VB_Name = "Test__MiscRange"
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

'@TestMethod("MiscRange")
Private Sub Test_activeRowsDown()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim RangeStart As Range
    Dim RangeTest As Range
    Dim Arr(2, 2) As Variant

    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "1"
    Arr(0, 1) = "1"
    Arr(0, 2) = "1"
    Arr(1, 0) = "1"
    Arr(1, 1) = "1"
    Arr(1, 2) = "1"
    Arr(2, 0) = "1"
    Arr(2, 1) = "1"
    Arr(2, 2) = "1"
    
    
    Set RangeStart = WB.ActiveSheet.Range("B4")
    Set RangeTest = ArrayToRange(Arr, RangeStart, True)
    
    'Assert:
    Assert.AreEqual CLng(3), ActiveRowsDown(WB.ActiveSheet.Range("B4"))
    Assert.AreEqual CLng(2), ActiveRowsDown(WB.ActiveSheet.Range("D5"))
    Assert.AreEqual CLng(1), ActiveRowsDown(WB.ActiveSheet.Range("C6"))

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub