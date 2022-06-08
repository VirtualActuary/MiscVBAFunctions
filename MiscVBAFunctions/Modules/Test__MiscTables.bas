Attribute VB_Name = "Test__MiscTables"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private WB As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)
    
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

'@TestMethod("MiscTables")
Private Sub TestHasLO_and_GetLO()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    
    'Assert:
    Assert.AreEqual True, HasLO("Table1", WB) ' Correct case
    Assert.AreEqual True, HasLO("taBLe1", WB) ' Should work without correct case
    
    Assert.AreEqual "Table1", GetLO("taBLe1", WB).Name ' Should get the correct name even with a different casing
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("MiscTables")
Private Sub TestListAllTables()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tables As Collection
    'Act:
    
    'Assert:
    Set Tables = GetAllTables(WB)
    
    Assert.IsTrue IsValueInCollection(Tables, "Table1")
    Assert.IsTrue IsValueInCollection(Tables, "NamedRange1")
    Assert.IsTrue IsValueInCollection(Tables, "SheetScopedNamedRange1")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTables")
Private Sub TestTableRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TableR As Range
    'Act:
    
    'Assert:
    Set TableR = TableRange("table1", WB)
    Assert.AreEqual "Column1", TableR.Cells(1).Value
    
    Set TableR = TableRange("NamedRange1", WB)
    Assert.AreEqual "Column1", TableR.Cells(1).Value
    
    Set TableR = TableRange("SheetScopedNamedRange1", WB)
    Assert.AreEqual "Column1", TableR.Cells(1).Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


