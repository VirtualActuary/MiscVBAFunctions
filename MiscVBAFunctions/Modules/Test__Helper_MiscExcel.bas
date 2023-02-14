Attribute VB_Name = "Test__Helper_MiscExcel"
Option Explicit

Function Test_ExcelBook()
    On Error GoTo TestFail1
    
    Dim WB1 As New Workbook
    Dim WB2 As New Workbook

    Set WB1 = ExcelBook(Path(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx"), True, True)
    Set WB2 = ExcelBook(Path(ThisWorkbook.Path, ".\test_data\MiscExcel\MiscExcel_added.xlsx"), False, True)
    WB1.Close False
    WB2.Close False

    Test_ExcelBook = True
    Exit Function
    
TestFail1:
    WB1.Close False
    WB2.Close False
    Test_ExcelBook = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile()
    On Error GoTo TestFail2
    
    Dim WB As New Workbook
    Set WB = ExcelBook()
    WB.Close False

    Test_ExcelBook_tempFile = True
    Exit Function

TestFail2:
    Test_ExcelBook_tempFile = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile_2()
    On Error GoTo TestFail3
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", False, False)
    WB.Close False

    Test_ExcelBook_tempFile_2 = True
    Exit Function

TestFail3:
    Test_ExcelBook_tempFile_2 = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile_fail()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail4
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", True, False)
    
    Test_ExcelBook_tempFile_fail = False
    Exit Function

TestFail4:
    If Err.Number = ExpectedError Then
        Test_ExcelBook_tempFile_fail = True
        Exit Function
    Else
        Test_ExcelBook_tempFile_fail = False
        Exit Function
    End If
End Function


Function Test_ExcelBook_tempFile_fail_2()
    Const ExpectedError As Long = -996
    On Error GoTo TestFail5
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", False, True)
    
    Test_ExcelBook_tempFile_fail_2 = False
    Exit Function
        
TestFail5:
    If Err.Number = ExpectedError Then
        Test_ExcelBook_tempFile_fail_2 = True
        Exit Function
    Else
        Test_ExcelBook_tempFile_fail_2 = False
        Exit Function
    End If
End Function


Function Test_fail_ExcelBook()
    Const ExpectedError As Long = -999
    On Error GoTo TestFail6
    
    Dim WB As New Workbook
    WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), True, True)
    
    Test_fail_ExcelBook = False
    Exit Function
TestFail6:
    If Err.Number = ExpectedError Then
        Test_fail_ExcelBook = True
        Exit Function
    Else
        Test_fail_ExcelBook = False
        Exit Function
    End If
End Function


Function Test_fail_ExcelBook_2()
    Const ExpectedError As Long = -998
    On Error GoTo TestFail7
    
    Dim WB As New Workbook
    WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), False, True)
    Test_fail_ExcelBook_2 = False
    Exit Function

TestFail7:
    If Err.Number = ExpectedError Then
        Test_fail_ExcelBook_2 = True
        Exit Function
    Else
        Test_fail_ExcelBook_2 = False
        Exit Function
    End If
End Function


Function Test_OpenWorkbook()
    On Error GoTo TestFail8
    
    Dim WB As Workbook
    Dim Asd As String
    Asd = Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx")
    
    Set WB = OpenWorkbook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx"), False)
    WB.Close False
    
    Test_OpenWorkbook = True
    Exit Function
    
TestFail8:
    Test_OpenWorkbook = False
    Exit Function
End Function


Function Test_LastRow()
    Dim WB As Workbook
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)

    Dim Rows As Integer
    Rows = LastRow(WB.Sheets(1))
    
    Test_LastRow = 19 = Rows
    WB.Close False
End Function


Function Test_LastColumn()
    Dim WB As Workbook
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)

    Dim Column As Integer
    Column = LastColumn(WB.Sheets(1))

    Test_LastColumn = 14 = Column
    WB.Close False
End Function


Function Test_LastCell_1()
    Dim WB As Workbook
    Dim R1 As Range
    Dim Pass As Boolean
    Pass = True

    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)
    Set R1 = LastCell(WB.Sheets(1))

    Pass = 1 = CInt(R1.Count) = Pass
    Pass = 1 = CInt(R1.Rows.Count) = Pass
    Pass = 1 = CInt(R1.Columns.Count) = Pass
    Pass = "$N$19" = R1.Address = Pass
    Pass = 100 = CInt(R1.Value) = Pass
    Test_LastCell_1 = Pass
    WB.Close False
End Function


Function Test_LastCell_2()
    Dim WB As Workbook
    Dim R1 As Range
    Dim Pass As Boolean
    Pass = True
    
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges.xlsx"), True, False)
    WB.Sheets(1).Cells(4, 14).Value = 4
    Set R1 = LastCell(WB.Sheets(1))

    Pass = 1 = CInt(R1.Count) = Pass
    Pass = 1 = CInt(R1.Rows.Count) = Pass
    Pass = 1 = CInt(R1.Columns.Count) = Pass
    Pass = "" = R1.Value = Pass
    Pass = "$N$11" = R1.Address = Pass

    Test_LastCell_2 = Pass
    WB.Close False
End Function


Function Test_RelevantRange()
    Dim WB As Workbook
    Dim R1 As Range
    Dim Arr As Variant
    Dim Pass As Boolean
    Pass = True
    
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges.xlsx"), True, True)
    Set R1 = RelevantRange(WB.Sheets(1))
    Arr = R1.Value
    
    Pass = 99 = CInt(R1.Count)
    Pass = 11 = CInt(R1.Rows.Count)
    Pass = 9 = CInt(R1.Columns.Count)
    Pass = 1 = CInt(LBound(Arr, 1))
    Pass = 11 = CInt(UBound(Arr, 1))
    Pass = 1 = CInt(LBound(Arr, 2))
    Pass = 9 = CInt(UBound(Arr, 2))
    Pass = "$I$11" = R1.Range(Cells(UBound(Arr, 1), UBound(Arr, 2)), Cells(UBound(Arr, 1), UBound(Arr, 2))).Address

    Test_RelevantRange = Pass
    WB.Close False
End Function


Function Test_RelevantRange2()
    Dim WB As Workbook
    Dim R1 As Range
    Dim Arr As Variant
    
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges2.xlsx"), True, True)
    Set R1 = RelevantRange(WB.Sheets(1))

    If R1 Is Nothing Then
        Test_RelevantRange2 = True
    Else
        Test_RelevantRange2 = False
    End If

    WB.Close False
End Function


Function Test_RenameSheet_fail()
    Const ExpectedError As Long = 58
    On Error GoTo TestFail
    
    Dim WB As Workbook
    Set WB = ExcelBook("")
    
    WB.Worksheets.Add , WB.Worksheets(1)
    RenameSheet WB.Worksheets(1), "foo"
    RenameSheet WB.Worksheets(2), "foo", True

    Test_RenameSheet_fail = False
    WB.Close False
    Exit Function
TestFail:
    If Err.Number = ExpectedError Then
        Test_RenameSheet_fail = True
    Else
        Test_RenameSheet_fail = False
    End If
    
    WB.Close False
End Function


Function Test_AddWS()
    Test_AddWS = False
    On Error GoTo TestFail
    
    Dim WB As Workbook
    Dim WS As Worksheet

    Set WB = ExcelBook("")
    Set WS = AddWS("NewSheet", WB:=WB)
    
    Test_AddWS = True
    WB.Close False
    Exit Function
TestFail:
    Test_AddWS = False
End Function


Function Test_InsertColumns()
    Dim Pass As Boolean
    Pass = True
    
    Dim WB As Workbook
    Dim RangeStart As Range
    Dim RelevantR As Range
  
    Set WB = ExcelBook("")
    Set RangeStart = WB.ActiveSheet.Range("C4")
    
    WB.ActiveSheet.Range("D5") = "foo"
    Set RelevantR = RelevantRange(WB.ActiveSheet)

    Pass = 20 = CInt(RelevantR.Count) = True = Pass
    Pass = 5 = CInt(RelevantR.Rows.Count) = True = Pass
    Pass = 4 = CInt(RelevantR.Columns.Count) = True = Pass
    
    InsertColumns RangeStart
    
    Pass = 25 = CInt(RelevantR.Count) = True = Pass
    Pass = 5 = CInt(RelevantR.Rows.Count) = True = Pass
    Pass = 5 = CInt(RelevantR.Columns.Count) = True = Pass
    
    Test_InsertColumns = Pass
    WB.Close False
End Function
