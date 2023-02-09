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

    'Act + Assert:
    Dim Dicts As Collection
    
    ' Read all columns:
    Set Dicts = TableToDicts("ListObject1", WB)
    Assert.AreEqual 5, CInt(Dicts(2)("b"))
    Assert.AreEqual 5, CInt(Dicts(2)("B")) ' must be case insensitive
    
    Set Dicts = TableToDicts("ListObject1", WB, col("a", "C"))
    Assert.AreEqual True, hasKey(Dicts(1), "A") ' should contain A
    Assert.AreNotEqual True, hasKey(Dicts(1), "b") ' should not contain b

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
    Dim Dicts As Collection
    
    'Act + Assert:
    ' Read all columns:
    Set Dicts = TableToDicts("NamedRange1", WB)
    Assert.AreEqual 5, CInt(Dicts(2)("b"))
    Assert.AreEqual 5, CInt(Dicts(2)("B")) ' must be case insensitive
    
    Set Dicts = TableToDicts("NamedRange1", WB, col("a", "C"))
    Assert.AreEqual True, hasKey(Dicts(1), "A") ' should contain A
    Assert.AreNotEqual True, hasKey(Dicts(1), "b") ' should not contain b
    
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
    Dim Dicts As Collection
    
    'Act + Assert:

    ' Read all columns:
    Set Dicts = TableToDicts("ListObject2", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    Set Dicts = TableToDicts("NamedRange2", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
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
    Dim Dicts As Collection
    
    'Act + Assert:
     
    ' Read all columns:
    Set Dicts = TableToDicts("ListObject3", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)
    
    Set Dicts = TableToDicts("NamedRange3", WB)
    Assert.AreEqual 0, CInt(Dicts.Count)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("MiscTableToDicts")
Private Sub TestGetTableRowIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Table As Collection
    
    'Act:
    Set Table = col(dicti("a", 1, "b", 2), dicti("a", 3, "b", 4), dicti("a", "foo", "b", "bar"), dicti("a", "Baz", "b", "Bla"))
    
    'Assert:
    Assert.AreEqual CLng(2), GetTableRowIndex(Table, col("a", "b"), col(3, 4))
    Assert.AreEqual CLng(3), GetTableRowIndex(Table, col("a", "b"), col("foo", "bar"))
    Assert.AreEqual CLng(3), GetTableRowIndex(Table, col("a", "b"), col("FoO", "BAr"))
    
    
    Dim IndexTest As Variant
    IndexTest = 999
    On Error GoTo NoFind
        ' this should throw an error as no match of the same index should be found
        IndexTest = GetTableRowIndex(Table, col("a", "b"), col("baz", "bla"), IgnoreCaseValues:=False)
NoFind:
    On Error GoTo TestFail
    Assert.AreEqual CLng(999), CLng(IndexTest)
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTableToDicts")
Private Sub TestTableLookupValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Table As Collection
    
    'Act:
    Set Table = col(dicti("a", 1, "b", 2, "c", 5), dicti("a", 3, "b", 4, "c", 6), dicti("a", "foo", "b", "bar"))
    
    'Assert:
    ' look for the value in column 'c' where column 'a' = 3 and column 'b' = 4:
    Assert.AreEqual CLng(6), CLng(TableLookupValue(Table, col("a", "b"), col(3, 4), "c"))
    ' Also test for when no lookup is found and default is given:
    Assert.AreEqual "foo", TableLookupValue(Table, col("a", "b"), col(3, 400), "c", "foo")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTableToDicts")
Private Sub TestTableToDictsLogSource()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dicts As Collection
    Dim Source As Dictionary
    
    'Act:
 
    Set Dicts = TableToDictsLogSource("ListObject1", WB)
    Set Source = dictget(Dicts(2), "__source__")
    
    'Assert:
    Assert.AreEqual "ListObject1", dictget(Source, "table")
    Assert.AreEqual CLng(2), dictget(Source, "rowindex")
    Assert.AreEqual "MiscTableToDicts.xlsx", dictget(Source, "workbook").Name
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTableToDicts")
Private Sub TestGetTableRowRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dicts As Collection
    Dim Source As Dictionary
    Dim R As Range
    
    'Act:
    ' Test list object:
    Set R = GetTableRowRange("ListObject1", col("a", "b"), col(4, 5), WB)
    Assert.AreEqual "$B$6:$D$6", R.Address
    
    Set R = GetTableRowRange("NamedRange1", col("a", "b"), col(4, 5), WB)
    Assert.AreEqual "$G$6:$I$6", R.Address

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTableToDicts")
Private Sub Test_TableDictToArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim col1 As Collection
    Dim Arr() As Variant
    
    'Act:
    Set col1 = col(Dict("a", 1, "b", 2), Dict("b", 11, "a", 12))
    Arr = TableDictToArray(col1)

    'Assert:
    Assert.AreEqual "a", Arr(0, 0)
    Assert.AreEqual "b", Arr(0, 1)
    Assert.AreEqual 1, Arr(1, 0)
    Assert.AreEqual 2, Arr(1, 1)
    Assert.AreEqual 12, Arr(2, 0)
    Assert.AreEqual 11, Arr(2, 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscTableToDicts")
Private Sub Test_TableDictToArray_fail_1()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail
    
    'Arrange:
    Dim col1 As Collection
    Dim Arr() As Variant
    
    'Act:
    Set col1 = col(Dict("a", 1, "b", 2), Dict("b", 11, "a", 12, "c", 3))
    Arr = TableDictToArray(col1)



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

'@TestMethod("MiscTableToDicts")
Private Sub Test_TableDictToArray_fail_2()
    Const ExpectedError As Long = -996
    On Error GoTo TestFail
    
    Dim col1 As Collection
    Dim Arr() As Variant
    
    'Act:
    Set col1 = col(Dict("a", 1, "b", 2), Dict("b", 11, "c", 3))
    Arr = TableDictToArray(col1)

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
