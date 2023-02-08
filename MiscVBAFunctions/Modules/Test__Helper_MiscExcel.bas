Attribute VB_Name = "Test__Helper_MiscExcel"
Option Explicit

Function Test_ExcelBook()
    On Error GoTo TestFail
    
    Dim WB1 As New Workbook
    Dim WB2 As New Workbook

    Set WB1 = ExcelBook(Path(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx"), True, True)
    Set WB2 = ExcelBook(Path(ThisWorkbook.Path, ".\test_data\MiscExcel\MiscExcel_added.xlsx"), False, True)
    WB1.Close False
    WB2.Close False

    Test_ExcelBook = True
    Exit Function
    
TestFail:
    WB1.Close False
    WB2.Close False
    Test_ExcelBook = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile()
    On Error GoTo Fail
    
    Dim WB As New Workbook
    Set WB = ExcelBook()
    WB.Close False

    Test_ExcelBook_tempFile = True
    Exit Function

Fail:
    Test_ExcelBook_tempFile = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile_2()
    On Error GoTo TestFail
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", False, False)
    WB.Close False

    Test_ExcelBook_tempFile_2 = True
    Exit Function

TestFail:
    Test_ExcelBook_tempFile_2 = False
    Exit Function
End Function


Function Test_ExcelBook_tempFile_fail()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", True, False)
    
    Test_ExcelBook_tempFile_fail = False
    Exit Function

TestFail:
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
    On Error GoTo TestFail
    
    Dim WB As New Workbook
    Set WB = ExcelBook("", False, True)
    
    Test_ExcelBook_tempFile_fail_2 = False
    Exit Function
        
TestFail:
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
    On Error GoTo TestFail
    
    Dim WB As New Workbook
    WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), True, True)
    
    Test_fail_ExcelBook = False
    Exit Function
TestFail:
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
    On Error GoTo TestFail
    
    Dim WB As New Workbook
    WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), False, True)
    Test_fail_ExcelBook_2 = False
    Exit Function

TestFail:
    If Err.Number = ExpectedError Then
        Test_fail_ExcelBook_2 = True
        Exit Function
    Else
        Test_fail_ExcelBook_2 = False
        Exit Function
    End If
End Function


Function Test_OpenWorkbook()
    On Error GoTo TestFail
    
    Dim WB As Workbook
    Dim Asd As String
    Asd = fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx")
    
    Set WB = OpenWorkbook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\MiscExcel.xlsx"), False)
    WB.Close False
    
    Test_OpenWorkbook = True
    Exit Function
    
TestFail:
    Test_OpenWorkbook = False
    Exit Function
End Function


Function Test_LastRow()
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)

    Dim Rows As Integer
    Rows = LastRow(WB.Sheets(1))
    
    Test_LastRow = 19 = Rows
    WB.Close False
End Function


Function Test_LastColumn()
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)

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

    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscTables\MiscTablesTests.xlsx"), True, True)
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
    
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges.xlsx"), True, False)
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
    
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges.xlsx"), True, True)
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
    
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "test_data\MiscExcel\ranges2.xlsx"), True, True)
    Set R1 = RelevantRange(WB.Sheets(1))

    If R1 Is Nothing Then
        Test_RelevantRange2 = True
    Else
        Test_RelevantRange2 = False
    End If

    WB.Close False
End Function
