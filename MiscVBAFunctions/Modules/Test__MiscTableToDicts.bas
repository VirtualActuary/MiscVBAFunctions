Attribute VB_Name = "Test__MiscTableToDicts"
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
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "tests\MiscTableToDicts\MiscTableToDicts.xlsx"), True, True)
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

'@TestMethod("MiscTableToDicts")
Private Sub TestListObjectsToDicts()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Dim Dicts As Collection
    
    ' Read all columns:
    Set Dicts = TableToDicts("ListObject1", WB)
    Assert.AreEqual CInt(Dicts(2)("b")), 5
    Assert.AreEqual CInt(Dicts(2)("B")), 5 ' must be case insensitive
    
    Set Dicts = TableToDicts("ListObject1", WB, col("a", "C"))
    Assert.AreEqual True, hasKey(Dicts(1), "A") ' should contain A
    Assert.AreNotEqual True, hasKey(Dicts(1), "b") ' should not contain b
    
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTableToDicts")
Private Sub TestNamedRangeToDicts()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Dim Dicts As Collection
    
    ' Read all columns:
    Set Dicts = TableToDicts("NamedRange1", WB)
    Assert.AreEqual 5, CInt(Dicts(2)("b"))
    Assert.AreEqual 5, CInt(Dicts(2)("B")) ' must be case insensitive
    
    Set Dicts = TableToDicts("NamedRange1", WB, col("a", "C"))
    Assert.AreEqual True, hasKey(Dicts(1), "A") ' should contain A
    Assert.AreNotEqual True, hasKey(Dicts(1), "b") ' should not contain b
    
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTableToDicts")
Private Sub TestEmptyTablesToDicts()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Dim Dicts As Collection
    
    ' Read all columns:
    Set Dicts = TableToDicts("ListObject2", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    Set Dicts = TableToDicts("NamedRange2", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTableToDicts")
Private Sub TestEmpty1ColumnTablesToDicts()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Dim Dicts As Collection
    
    ' Read all columns:
    Set Dicts = TableToDicts("ListObject3", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    Set Dicts = TableToDicts("NamedRange3", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
