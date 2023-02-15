Attribute VB_Name = "Test__Helper_MiscTables"
Option Explicit

Function TestHasLO_and_GetLO(WB As Workbook)
    Dim Pass As Boolean
    Pass = True
    
    Pass = True = HasLO("Table1", WB) = Pass = True ' Correct case
    Pass = True = HasLO("taBLe1", WB) = Pass = True ' Should work without correct case

    Pass = "Table1" = GetLO("taBLe1", WB).Name = Pass = True ' Should get the correct name even with a different casing

    TestHasLO_and_GetLO = Pass
End Function


Function TestTableRange(WB As Workbook)
    Dim TableR As Range
    Dim Pass As Boolean
    Pass = True

    Set TableR = TableRange("table1", WB)
    Pass = "Column1" = TableR.Cells(1).Value = Pass = True

    Set TableR = TableRange("NamedRange1", WB)
    Pass = "Column1" = TableR.Cells(1).Value = Pass = True

    Set TableR = TableRange("SheetScopedNamedRange1", WB)
    Pass = "Column1" = TableR.Cells(1).Value = Pass = True

    TestTableRange = Pass
End Function


Function Test_CopyTable(WB As Workbook, WB2 As Workbook)
    Dim Pass As Boolean
    Pass = True

    Dim LO As ListObject
    Dim LOEntries As Range

    CopyTable "TableForCopy", WB2.Worksheets(1).Cells(5, 3), , WB, True
    Set LO = GetLO("TableForCopy", WB2)
    Set LOEntries = LO.DataBodyRange

    Pass = HasLO("TableForCopy", WB2) = Pass = True
    
    Pass = "General" = LOEntries(1).NumberFormat = Pass = True
    Pass = "0.00%" = LOEntries(2).NumberFormat = Pass = True
    Pass = "$#,##0.00" = LOEntries(3).NumberFormat = Pass = True
    Pass = "General" = LOEntries(4).NumberFormat = Pass = True
    Pass = "General" = LOEntries(5).NumberFormat = Pass = True
    Pass = "m/d/yyyy" = LOEntries(6).NumberFormat = Pass = True

    Pass = "=[foo]" = LOEntries(1, 1).Value = Pass = True
    Pass = 12 = CInt(LOEntries(1, 2).Value) = Pass = True
    Pass = 21 = CInt(LOEntries(2, 1).Value) = Pass = True
    Pass = 22 = CInt(LOEntries(2, 2).Value) = Pass = True
    Pass = "Hello" = LOEntries(3, 1).Value = Pass = True
    Pass = 32 = CInt(LOEntries(3, 2).Value) = Pass = True
    Pass = CVErr(xlErrName) = LOEntries(4, 1).Value = Pass = True
    Pass = CVErr(xlErrNA) = LOEntries(4, 2).Value = Pass = True
    Pass = "=foo" = LOEntries(5, 1).Value = Pass = True
    
    Test_CopyTable = Pass
End Function


Function Test_TableColumnToCollection()
    Dim Pass As Boolean
    Pass = True

    Dim Col1 As Collection
    Dim Col2 As Collection

    Set Col1 = Col(Dict("a", 1, "b", 2), Dict("a", 10, "b", 20))
    Set Col2 = TableColumnToCollection(Col1, "b")

    Pass = 2 = Col2(1) = Pass = True
    Pass = 20 = Col2(2) = Pass = True

    Test_TableColumnToCollection = Pass
End Function


Function Test_ResizeLO_1(SelectedTable As ListObject)
    Test_ResizeLO_1 = 3 = CInt(SelectedTable.ListRows.Count)
End Function


Function Test_ResizeLO_2(SelectedTable As ListObject)
    Test_ResizeLO_2 = 0 = CInt(SelectedTable.ListRows.Count)
End Function


Function Test_ResizeLO_3(SelectedTable As ListObject)
    Test_ResizeLO_3 = 2 = CInt(SelectedTable.ListRows.Count)
End Function


Function Test_ResizeLO_4(SelectedTable As ListObject)
     Test_ResizeLO_4 = 1 = CInt(SelectedTable.ListRows.Count)
End Function


Function Test_GetTableColumnDataRange_fail(SelectedTable As ListObject)
    Const ExpectedError As Long = 32000
    On Error GoTo TestFail

    Dim R As Range
    Set R = GetTableColumnDataRange(SelectedTable, "NonExistingColumn")

    Test_GetTableColumnDataRange_fail = False
    Exit Function
TestFail:
    If Err.Number = ExpectedError Then
        Test_GetTableColumnDataRange_fail = True
        Exit Function
    Else
        Test_GetTableColumnDataRange_fail = False
        Exit Function
    End If
End Function


Function Test_GetTableRowNumberDataRange_fail(SelectedTable As ListObject)
    Const ExpectedError As Long = 32000
    On Error GoTo TestFail

    Dim R As Range
    Set R = GetTableRowNumberDataRange(SelectedTable, 20)

    Test_GetTableRowNumberDataRange_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_GetTableRowNumberDataRange_fail = True
    Else
        Test_GetTableRowNumberDataRange_fail = False
    End If
End Function


Function Test_GetTableRowRange1(R As Range)
    Test_GetTableRowRange1 = "$B$6:$D$6" = R.Address
End Function


Function Test_GetTableRowRange2(R As Range)
    Test_GetTableRowRange2 = "$G$6:$L$6" = R.Address
End Function

