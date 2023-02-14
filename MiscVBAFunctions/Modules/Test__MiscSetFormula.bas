Attribute VB_Name = "Test__MiscSetFormula"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
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

'@TestMethod("SetFormula")
Private Sub Test_SetFormula_1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook("")
    Dim Sht As Worksheet
    Set Sht = WB.Worksheets(1)
    Dim rng As Range
    Set rng = Sht.Range("B2")
    
    'Act:
    SetFormula rng, "=1+2.1"
    
    'Assert:
    Assert.AreEqual 3.1, rng.Value
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
