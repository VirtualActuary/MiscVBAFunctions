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

'@TestMethod("MiscTables")
Private Sub Test_TableColumnToArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim col1 As Collection
    Dim arr() As Variant
    'Act:
    Set col1 = col(dict("a", 1, "b", 2), dict("a", 10, "b", 20))
    arr = TableColumnToArray(col1, "b")
    
    'Assert:
    Assert.AreEqual 2, arr(0)
    Assert.AreEqual 20, arr(1)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_TableColumnToCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim col1 As Collection
    Dim col2 As Collection
    'Act:
    Set col1 = col(dict("a", 1, "b", 2), dict("a", 10, "b", 20))
    Set col2 = TableColumnToCollection(col1, "b")
    
    'Assert:
    Assert.AreEqual 2, col2(1)
    Assert.AreEqual 20, col2(2)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_ResizeLO_1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    
    'Act:
    Set SelectedTable = GetLO("Table1", WB)
    ResizeLO SelectedTable, 3

    'Assert:
    Assert.AreEqual 3, CInt(SelectedTable.ListRows.Count)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_ResizeLO_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    
    'Act:
    Set SelectedTable = GetLO("Table1", WB)
    ResizeLO SelectedTable, 0

    'Assert:
    Assert.AreEqual 0, CInt(SelectedTable.ListRows.Count)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_ResizeLO_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    
    'Act:
    Set SelectedTable = GetLO("Table1", WB)
    ResizeLO SelectedTable, 5
    ResizeLO SelectedTable, 2

    'Assert:
    Assert.AreEqual 2, CInt(SelectedTable.ListRows.Count)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_ResizeLO_4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    
    'Act:
    Set SelectedTable = GetLO("Table1", WB)
    ResizeLO SelectedTable, 0
    ResizeLO SelectedTable, 1

    'Assert:
    Assert.AreEqual 1, CInt(SelectedTable.ListRows.Count)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_GetTableColumnDataRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    Dim r As Range
    Dim arr() As Variant
    
    'Act:
    Set SelectedTable = GetLO("table2", WB)
    Set r = GetTableColumnDataRange(SelectedTable, "Column2")
    arr = r.Value
    'Assert:
    Assert.AreEqual 12, CInt(arr(1, 1))
    Assert.AreEqual 22, CInt(arr(2, 1))
    Assert.AreEqual 32, CInt(arr(3, 1))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_GetTableColumnDataRange_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    Dim r As Range
    Dim arr() As Variant
    
    'Act:
    Set SelectedTable = GetLO("table2", WB)
    ResizeLO SelectedTable, 0
    Set r = GetTableColumnDataRange(SelectedTable, "Column2")
    
'    'Assert:
    If r Is Nothing Then
        Assert.Succeed
    Else
        Assert.Fail
    End If

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_GetTableColumnDataRange_fail()
    Const ExpectedError As Long = 32000
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    Dim r As Range
    
    'Act:
    Set SelectedTable = GetLO("table2", WB)
    Set r = GetTableColumnDataRange(SelectedTable, "NonExistingColumn")

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscTables")
Private Sub Test_GetTableRowNumberDataRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    Dim r As Range
    Dim arr() As Variant
    
    'Act:
    Set SelectedTable = GetLO("table2", WB)
    Set r = GetTableRowNumberDataRange(SelectedTable, 2)
    arr = r.Value
    'Assert:
    Assert.AreEqual 21, CInt(arr(1, 1))
    Assert.AreEqual 22, CInt(arr(1, 2))
    Assert.AreEqual 23, CInt(arr(1, 3))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTables")
Private Sub Test_GetTableRowNumberDataRange_fail()
    Const ExpectedError As Long = 32000
    On Error GoTo TestFail
    
    'Arrange:
    Dim SelectedTable As ListObject
    Dim r As Range
    
    'Act:
    Set SelectedTable = GetLO("table2", WB)
    Set r = GetTableRowNumberDataRange(SelectedTable, 20)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("MiscTables")
Private Sub Test_GetTableRowRange()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Dim Dicts As Collection
    Dim Source As Dictionary
    Dim WB2 As Workbook
    Set WB2 = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "tests\MiscTableToDicts\MiscTableToDicts.xlsx"), True, True)
    ' Test list object:
    Dim r As Range
    Set r = GetTableRowRange("ListObject1", col("a", "b"), col(4, 5), WB2)
    Assert.AreEqual "$B$6:$D$6", r.Address
    
    Set r = GetTableRowRange("NamedRange1", col("a", "b"), col(4, 5), WB2)
    Assert.AreEqual "$G$6:$I$6", r.Address

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

