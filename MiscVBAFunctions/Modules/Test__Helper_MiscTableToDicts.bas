Attribute VB_Name = "Test__Helper_MiscTableToDicts"
Option Explicit

Function TestListObjectsToDicts1(Dicts)
    Dim Pass As Boolean
    Pass = True
    
    Pass = 5 = CInt(Dicts(2)("b")) = Pass = True
    Pass = 5 = CInt(Dicts(2)("B")) = Pass = True ' must be case insensitive
    
    TestListObjectsToDicts1 = Pass
End Function


Function TestListObjectsToDicts2(Dicts)
    Dim Pass As Boolean
    Pass = True
    
    Pass = HasKey(Dicts(1), "A") = Pass = True ' should contain A
    Pass = (Not HasKey(Dicts(1), "b")) = Pass = True ' should not contain b
    
    TestListObjectsToDicts2 = Pass
End Function


Function TestNamedRangeToDicts1(Dicts)
    Dim Pass As Boolean
    Pass = True

    Pass = 5 = CInt(Dicts(2)("b")) = Pass = True
    Pass = 5 = CInt(Dicts(2)("B")) = Pass = True ' must be case insensitive
    
    TestNamedRangeToDicts1 = Pass
End Function


Function TestNamedRangeToDicts2(Dicts)
    Dim Pass As Boolean
    Pass = True

    Pass = True = HasKey(Dicts(1), "A") = Pass = True ' should contain A
    Pass = True = (Not HasKey(Dicts(1), "b")) = Pass = True ' should not contain b
    
    TestNamedRangeToDicts2 = Pass
End Function


Function TestEmptyTablesToDicts1(WB As Workbook)
    Dim Dicts As Collection
    Set Dicts = TableToDicts("ListObject2", WB)
    TestEmptyTablesToDicts1 = 0 = CInt(Dicts.Count)
End Function


Function TestEmptyTablesToDicts2(WB As Workbook)
    Dim Dicts As Collection
    Set Dicts = TableToDicts("NamedRange2", WB)
    TestEmptyTablesToDicts2 = 0 = CInt(Dicts.Count)
End Function


Function TestEmpty1ColumnTablesToDicts1(WB As Workbook)
    Dim Dicts As Collection
    Set Dicts = TableToDicts("ListObject3", WB)
    TestEmpty1ColumnTablesToDicts1 = 0 = CInt(Dicts.Count)
End Function


Function TestEmpty1ColumnTablesToDicts2(WB As Workbook)
    Dim Dicts As Collection
    Set Dicts = TableToDicts("NamedRange3", WB)
    TestEmpty1ColumnTablesToDicts2 = 0 = CInt(Dicts.Count)
End Function


Function TestGetTableRowIndex1()
    Dim Pass As Boolean
    Pass = True
    Dim Table As Collection
    Set Table = Col(DictI("a", 1, "b", 2), DictI("a", 3, "b", 4), DictI("a", "foo", "b", "bar"), DictI("a", "Baz", "b", "Bla"))

    Pass = CLng(2) = GetTableRowIndex(Table, Col("a", "b"), Col(3, 4)) = Pass = True
    Pass = CLng(3) = GetTableRowIndex(Table, Col("a", "b"), Col("foo", "bar")) = Pass = True
    Pass = CLng(3) = GetTableRowIndex(Table, Col("a", "b"), Col("FoO", "BAr")) = Pass = True
    
    TestGetTableRowIndex1 = Pass
End Function


Function TestGetTableRowIndex2()
    Dim Table As Collection
    Set Table = Col(DictI("a", 1, "b", 2), DictI("a", 3, "b", 4), DictI("a", "foo", "b", "bar"), DictI("a", "Baz", "b", "Bla"))

    Dim IndexTest As Variant
    IndexTest = 999
    On Error GoTo NoFind
    ' this should throw an error as no match of the same index should be found
    IndexTest = GetTableRowIndex(Table, Col("a", "b"), Col("baz", "bla"), IgnoreCaseValues:=False)
    
    TestGetTableRowIndex2 = False
    Exit Function
NoFind:
    TestGetTableRowIndex2 = CLng(999) = CLng(IndexTest)
End Function


Function TestTableToDictsLogSource(WB As Workbook)
    Dim Pass As Boolean
    Pass = True
    
    Dim Dicts As Collection
    Dim Source As Dictionary

    Set Dicts = TableToDictsLogSource("ListObject1", WB)
    Set Source = Dictget(Dicts(2), "__source__")
    
    Pass = "ListObject1" = Dictget(Source, "table") = Pass = True
    Pass = CLng(2) = Dictget(Source, "rowindex") = Pass = True
    Pass = "MiscTableToDicts.xlsx" = Dictget(Source, "workbook").Name = Pass = True
    
    TestTableToDictsLogSource = Pass
End Function


Function TestGetTableRowRange1(WB As Workbook)
    Dim R As Range
    Set R = GetTableRowRange("ListObject1", Col("a", "b"), Col(4, 5), WB)
    TestGetTableRowRange1 = "$B$6:$D$6" = R.Address
End Function


Function TestGetTableRowRange2(WB As Workbook)
    Dim R As Range
    Set R = GetTableRowRange("NamedRange1", Col("a", "b"), Col(4, 5), WB)
    TestGetTableRowRange2 = "$G$6:$L$6" = R.Address
End Function


Function Test_TableDictToArray_fail_1()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail1
    
    Dim Col1 As Collection
    Dim Arr() As Variant
    
    Set Col1 = Col(Dict("a", 1, "b", 2), Dict("b", 11, "a", 12, "c", 3))
    Arr = TableDictToArray(Col1)

    Test_TableDictToArray_fail_1 = False
    Exit Function
    
TestFail1:
    If Err.Number = ExpectedError Then
        Test_TableDictToArray_fail_1 = True
    Exit Function
    Else
        Test_TableDictToArray_fail_1 = False
    Exit Function
    End If
End Function


Function Test_TableDictToArray_fail_2()
    Const ExpectedError As Long = -996
    On Error GoTo TestFail2
    
    Dim Col1 As Collection
    Dim Arr() As Variant
    
    Set Col1 = Col(Dict("a", 1, "b", 2), Dict("b", 11, "c", 3))
    Arr = TableDictToArray(Col1)
    
    Test_TableDictToArray_fail_2 = False
    Exit Function
TestFail2:
    If Err.Number = ExpectedError Then
        Test_TableDictToArray_fail_2 = True
        Exit Function
    Else
        Test_TableDictToArray_fail_2 = False
    Exit Function
    End If
End Function

