Attribute VB_Name = "Test__MiscFreezePanes"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private WS As Worksheet
Private WB As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    'Dim WS As Worksheet
    'Dim WB As Workbook
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscFreezePanes\MiscFreezePanes.xlsx"), True, True)
    Set WS = WB.Worksheets(1)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    WB.Close False
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("MiscFreezePanes")
Private Sub Test_FreezePanes()
    On Error GoTo TestFail
    
    'Arrange:
    'Dim WS As Worksheet
    Dim WB As Workbook
    'Act:
    
    FreezePanes WS.Range("D6")
    
    'Assert:
    With Application.Windows(WS.Parent.Name)
        Assert.IsTrue .FreezePanes
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscFreezePanes")
Private Sub Test_UnFreezePanes()
    On Error GoTo TestFail
    
    'Arrange:
    'Dim WS As Worksheet
    'Dim WB As Workbook
    
    
    'Act:
    'Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscFreezePanes\MiscFreezePanes.xlsx"), True, True)
    'Set WS = WB.Worksheets(1)
    FreezePanes WS.Range("D6")
    UnFreezePanes WS
    
    'Assert:
    With Application.Windows(WS.Parent.Name)
        Assert.IsFalse .FreezePanes
    End With
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

