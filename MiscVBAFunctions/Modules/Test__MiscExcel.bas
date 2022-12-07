Attribute VB_Name = "Test__MiscExcel"
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

'@TestMethod("MiscExcel")
Private Sub Test_ExcelBook()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB1 As New Workbook
    Dim WB2 As New Workbook
    
    'Act:
    Set WB1 = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel.xlsx"), True, True)
    Set WB2 = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel_added.xlsx"), False, True)
    WB1.Close False
    WB2.Close False

    'Assert:
    Assert.Succeed

TestExit:
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscExcel")
Private Sub Test_ExcelBook_tempFile()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    'Act:
    Set WB = ExcelBook()
    WB.Close False
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscExcel")
Private Sub Test_ExcelBook_tempFile_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    'Act:
    Set WB = ExcelBook("", False, False)
    WB.Close False
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_ExcelBook_tempFile_fail()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    'Act:
    Set WB = ExcelBook("", True, False)

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

'@TestMethod("MiscExcel")
Private Sub Test_ExcelBook_tempFile_fail_2()
    Const ExpectedError As Long = -996
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    'Act:
    Set WB = ExcelBook("", False, True)

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




'@TestMethod("MiscExcel")
Private Sub Test_fail_ExcelBook()
    Const ExpectedError As Long = -999
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    
    'Act:
    WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), True, True)
    
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

'@TestMethod("MiscExcel")
Private Sub Test_fail_ExcelBook_2()
    Const ExpectedError As Long = -998
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    
    'Act:
    WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\nonExistingFile.xlsx"), False, True)
    
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


'@TestMethod("MiscExcel")
Private Sub Test_OpenWorkbook()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    'Act:
    Set WB = OpenWorkbook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel.xlsx"), False)
    WB.Close False
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscExcel")
Private Sub Test_LastRow()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)

    'Act:
    Dim Rows As Integer
    
    Rows = LastRow(WB.Sheets(1))
    'Assert:
    Assert.AreEqual 19, Rows

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_LastColumn()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)

    Dim Column As Integer
    
    Column = LastColumn(WB.Sheets(1))
    'Assert:
    Assert.AreEqual 14, Column

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscExcel")
Private Sub Test_LastCell_1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim R1 As Range
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)
    Set R1 = LastCell(WB.Sheets(1))

    'Assert:
    Assert.AreEqual 1, CInt(R1.Count)
    Assert.AreEqual 1, CInt(R1.Rows.Count)
    Assert.AreEqual 1, CInt(R1.Columns.Count)
    Assert.AreEqual "$N$19", R1.Address
    Assert.AreEqual 100, CInt(R1.Value)

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_LastCell_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim R1 As Range
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges.xlsx"), True, False)
    WB.Sheets(1).Cells(4, 14).Value = 4
    Set R1 = LastCell(WB.Sheets(1))

    'Assert:
    Assert.AreEqual 1, CInt(R1.Count)
    Assert.AreEqual 1, CInt(R1.Rows.Count)
    Assert.AreEqual 1, CInt(R1.Columns.Count)
    Assert.AreEqual "", R1.Value
    Assert.AreEqual "$N$11", R1.Address

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_RelevantRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim R1 As Range
    Dim Arr As Variant
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges.xlsx"), True, True)
    Set R1 = RelevantRange(WB.Sheets(1))
    Arr = R1.Value
    
    'Assert:
    Assert.AreEqual 99, CInt(R1.Count)
    Assert.AreEqual 11, CInt(R1.Rows.Count)
    Assert.AreEqual 9, CInt(R1.Columns.Count)
    Assert.AreEqual 1, CInt(LBound(Arr, 1))
    Assert.AreEqual 11, CInt(UBound(Arr, 1))
    Assert.AreEqual 1, CInt(LBound(Arr, 2))
    Assert.AreEqual 9, CInt(UBound(Arr, 2))
    Assert.AreEqual "$I$11", R1.Range(Cells(UBound(Arr, 1), UBound(Arr, 2)), Cells(UBound(Arr, 1), UBound(Arr, 2))).Address

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_RelevantRange2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim R1 As Range
    Dim Arr As Variant
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges2.xlsx"), True, True)
    Set R1 = RelevantRange(WB.Sheets(1))

    'Assert:
    If R1 Is Nothing Then
        Assert.Succeed
    Else
        Assert.Fail
    End If

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
                                                                        

'@TestMethod("MiscExcel")
Private Sub Test_SanitiseExcelName()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Assert.AreEqual "_1", SanitiseExcelName("1")
    Assert.AreEqual "a_b", SanitiseExcelName("a b")
    Assert.AreEqual "______________________________", SanitiseExcelName("- /*+=^!@#$%&?`~:;[](){}""'|,<>")
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_VbaLocked()
    On Error GoTo TestFail
 
    'Assert:
    Assert.AreEqual ThisWorkbook.VBProject.Protection <> vbext_ProjectProtection.vbext_pp_none, VbaLocked()

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_RenameSheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook

    'Act:
    Set WB = ExcelBook("")
    
    RenameSheet "foo", WB.Worksheets(1)

    'Assert:
    Assert.AreEqual "foo", WB.Worksheets(1).Name

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_RenameSheet_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook

    'Act:
    Set WB = ExcelBook("")
    
    WB.Worksheets.Add , WB.Worksheets(1)
    RenameSheet "foo", WB.Worksheets(1)
    RenameSheet "foo", WB.Worksheets(2)

    'Assert:
    Assert.AreEqual "foo", WB.Worksheets(1).Name

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_RenameSheet_fail()
    Const ExpectedError As Long = 58
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook

    'Act:
    Set WB = ExcelBook("")
    
    WB.Worksheets.Add , WB.Worksheets(1)
    RenameSheet "foo", WB.Worksheets(1)
    RenameSheet "foo", WB.Worksheets(2), True

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscExcel")
Private Sub Test_AddWS()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim WS As Worksheet
    
    'Act:
    Set WB = ExcelBook("")
    Set WS = AddWS("NewSheet", WB:=WB)
    
    'Assert:
    Assert.Succeed

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
