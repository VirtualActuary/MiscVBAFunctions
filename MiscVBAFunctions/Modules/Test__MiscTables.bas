Attribute VB_Name = "Test__MiscTables"
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

'@TestMethod("MiscTables")
Private Sub TestHasLO_and_GetLO()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook, LO As ListObject
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)
    
    'Assert:
    Assert.AreEqual True, HasLO("Table1", WB) ' Correct case
    Assert.AreEqual True, HasLO("taBLe1", WB) ' Should work without correct case
    
    Assert.AreEqual "Table1", GetLO("taBLe1", WB).Name ' Should get the correct name even with a different casing
    
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
