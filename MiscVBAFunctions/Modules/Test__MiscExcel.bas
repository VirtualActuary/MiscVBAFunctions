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





