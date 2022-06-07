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
    Dim rows As Integer
    
    rows = LastRow(WB.Sheets(1))
    'Assert:
    Assert.AreEqual 6, rows

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

    Dim column As Integer
    
    column = LastColumn(WB.Sheets(1))
    'Assert:
    Assert.AreEqual 9, column

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
    Dim r1 As Range
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscTables\MiscTablesTests.xlsx"), True, True)
    Set r1 = LastCell(WB.Sheets(1))

    'Assert:
    Assert.AreEqual 1, CInt(r1.Count)
    Assert.AreEqual 1, CInt(r1.rows.Count)
    Assert.AreEqual 1, CInt(r1.Columns.Count)
    Assert.AreEqual "$I$6", r1.Address
    Assert.AreEqual "bla", r1.Value

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
    Dim r1 As Range
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges.xlsx"), True, False)
    WB.Sheets(1).Cells(4, 14).Value = 4
    Set r1 = LastCell(WB.Sheets(1))

    'Assert:
    Assert.AreEqual 1, CInt(r1.Count)
    Assert.AreEqual 1, CInt(r1.rows.Count)
    Assert.AreEqual 1, CInt(r1.Columns.Count)
    Assert.AreEqual "", r1.Value
    Assert.AreEqual "$N$11", r1.Address

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
    Dim r1 As Range
    Dim arr As Variant
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges.xlsx"), True, True)
    Set r1 = RelevantRange(WB.Sheets(1))
    arr = r1.Value
    
    'Assert:
    Assert.AreEqual 99, CInt(r1.Count)
    Assert.AreEqual 11, CInt(r1.rows.Count)
    Assert.AreEqual 9, CInt(r1.Columns.Count)
    Assert.AreEqual 1, CInt(LBound(arr, 1))
    Assert.AreEqual 11, CInt(UBound(arr, 1))
    Assert.AreEqual 1, CInt(LBound(arr, 2))
    Assert.AreEqual 9, CInt(UBound(arr, 2))
    Assert.AreEqual "$I$11", r1.Range(Cells(UBound(arr, 1), UBound(arr, 2)), Cells(UBound(arr, 1), UBound(arr, 2))).Address

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
    Dim r1 As Range
    Dim arr As Variant
    
    'Act:
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\ranges2.xlsx"), True, True)
    Set r1 = RelevantRange(WB.Sheets(1))

    'Assert:
    If r1 Is Nothing Then
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
