Attribute VB_Name = "Test__MiscArray"
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

'@TestMethod("MiscArray")
Private Sub Test_EnsureDotSeparatorTransformation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim I As Long, J As Long
    Dim Arr(2, 2)
    
    Dim Arr2(3)
    
    'Act:
    Arr(0, 0) = 100.2: Arr(0, 1) = 1.9
    Arr(1, 0) = 2.1: Arr(1, 1) = 2.2
    EnsureDotSeparatorTransformation Arr

    Arr2(0) = 1.2: Arr2(1) = 2.1: Arr2(2) = 3.8
    EnsureDotSeparatorTransformation Arr2
    
    
    'Assert:
    Assert.AreEqual "100.2", Arr(0, 0)
    Assert.AreEqual "1.9", Arr(0, 1)
    Assert.AreEqual "2.1", Arr(1, 0)
    Assert.AreEqual "2.2", Arr(1, 1)
    
    Assert.AreEqual "1.2", Arr2(0)
    Assert.AreEqual "2.1", Arr2(1)
    Assert.AreEqual "3.8", Arr2(2)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ErrorToNullStringTransformation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(2, 2)
    Dim arrSecond(3)
    
    'Act:
    Arr(0, 0) = 100.2: Arr(0, 1) = CVErr(xlErrName)
    Arr(1, 0) = 2.1: Arr(1, 1) = CVErr(xlErrNA)
    ErrorToNullStringTransformation Arr

    arrSecond(0) = 1.2: arrSecond(1) = CVErr(xlErrRef): arrSecond(2) = 3.8
    ErrorToNullStringTransformation arrSecond


    'Assert:
    Assert.AreEqual 100.2, Arr(0, 0)
    Assert.AreEqual vbNullString, Arr(0, 1)
    Assert.AreEqual 2.1, Arr(1, 0)
    Assert.AreEqual vbNullString, Arr(1, 1)
    
    Assert.AreEqual 1.2, arrSecond(0)
    Assert.AreEqual vbNullString, arrSecond(1)
    Assert.AreEqual 3.8, arrSecond(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_DateToStringTransformation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(2, 2)
    Dim arrSecond(3)
    Dim arrThird(1)
    Dim arrFourth(1)
    Dim arrFifth(1)

    'Act:
    Arr(0, 0) = CDate("2021-1-2"): Arr(0, 1) = CDate("2021-01-28 10:2")
    Arr(1, 0) = 13: Arr(1, 1) = 2.5
    DateToStringTransformation Arr

    arrSecond(0) = 1.2: arrSecond(1) = 2.1: arrSecond(2) = CDate("2021-3-28 10:2:10")
    DateToStringTransformation arrSecond
    
    arrThird(0) = CDate("2021-01-28 10:2:10")
    arrFourth(0) = CDate("2021-01-28 10:2:10")
    arrFifth(0) = CDate("2021-01-28 10:2:10")
    
    'Assert:
    Assert.AreEqual "2021-01-02", Arr(0, 0)
    Assert.AreEqual "2021-01-28", Arr(0, 1)
    Assert.AreEqual 13, Arr(1, 0)
    Assert.AreEqual 2.5, Arr(1, 1)
    
    Assert.AreEqual 1.2, arrSecond(0)
    Assert.AreEqual 2.1, arrSecond(1)
    Assert.AreEqual "2021-03-28", arrSecond(2)
    
    Assert.AreEqual "2021-01", DateToStringTransformation(arrThird, "yyyy-mm")(0)
    Assert.AreEqual "2021/01/28", DateToStringTransformation(arrFourth, "yyyy/mm/dd")(0)
    Assert.AreEqual "2021-01-28 10:02:10", DateToStringTransformation(arrFifth, "yyyy-mm-dd hh:mm:ss")(0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(3) As Variant
    Dim Col1 As Collection
    'Act:
    Arr(0) = 10
    Arr(1) = 11
    Arr(2) = 12
    Arr(3) = 13
    Set Col1 = ArrayToCollection(Arr)
    'Assert:
    Assert.AreEqual 10, CInt(Col1(1))
    Assert.AreEqual 11, CInt(Col1(2))
    Assert.AreEqual 12, CInt(Col1(3))
    Assert.AreEqual 13, CInt(Col1(4))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 2) As Variant
    Dim RangeObj As Range
    Dim RangeOutput As Range
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "col1"
    Arr(0, 1) = "col2"
    Arr(0, 2) = "col3"
    Arr(1, 0) = "=[d]"
    Arr(1, 1) = "=d"
    Arr(1, 2) = 1
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set RangeOutput = ArrayToRange(Arr, RangeObj, True)

    'Assert:
    Assert.AreEqual 6, CInt(RangeOutput.Count)
    Assert.AreEqual 2, CInt(RangeOutput.Column)
    Assert.AreEqual 4, CInt(RangeOutput.Row)
    Assert.AreEqual "col2", RangeOutput(1, 2).Value
    Assert.AreEqual 1, CInt(RangeOutput(2, 3).Value)
    

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToRange2dWithOneColumn()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 0) As Variant
    Dim RangeObj As Range
    Dim RangeOutput As Range
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "col1"
    Arr(1, 0) = "=[d]"
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set RangeOutput = ArrayToRange(Arr, RangeObj, True)

    'Assert:
    Assert.AreEqual 2, CInt(RangeOutput.Count)
    Assert.AreEqual 2, CInt(RangeOutput.Column)
    Assert.AreEqual 4, CInt(RangeOutput.Row)
    Assert.AreEqual "col1", RangeOutput(1, 1).Value
    Assert.AreEqual "=[d]", CStr(RangeOutput(2, 1).Value)
    

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToRange_FunkyHeaders()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 3) As Variant
    Dim RangeObj As Range
    Dim RangeOutput As Range
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "asdf"
    Arr(0, 1) = 1234
    Arr(0, 2) = "2022/11/02"
    Arr(0, 3) = False
    Arr(1, 0) = "a"
    Arr(1, 1) = "b"
    Arr(1, 2) = "c"
    Arr(1, 3) = "d"
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set RangeOutput = ArrayToRange(Arr, RangeObj, False, True)

    'Assert:
    Assert.AreEqual 8, CInt(RangeOutput.Count)
    Assert.AreEqual 4, CInt(RangeOutput.Columns.Count)
    Assert.AreEqual 2, CInt(RangeOutput.Rows.Count)
    Assert.AreEqual "asdf", RangeOutput.Cells(1, 1).Text
    Assert.AreEqual "1234", RangeOutput.Cells(1, 2).Text
    Assert.AreEqual "2022/11/02", RangeOutput.Cells(1, 3).Text
    Assert.AreEqual "FALSE", RangeOutput.Cells(1, 4).Text

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToNewTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 2) As Variant
    Dim RangeObj As Range
    Dim LO As ListObject
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "col1"
    Arr(0, 1) = "col2"
    Arr(0, 2) = "col3"
    Arr(1, 0) = "=[d]"
    Arr(1, 1) = "=d"
    Arr(1, 2) = 1
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, True)
    
    'Assert:
    Assert.AreEqual "TestTable", LO.Name
    Assert.AreEqual "col2", LO.Range(1, 2).Value
    Assert.AreEqual 1, CInt(LO.Range(2, 3).Value)
    
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToNewTable_FunkyHeaders()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 3) As Variant
    Dim RangeObj As Range
    Dim LO As ListObject
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "asdf"
    Arr(0, 1) = 1234
    Arr(0, 2) = "2022/11/02"
    Arr(0, 3) = False
    Arr(1, 0) = "a"
    Arr(1, 1) = "b"
    Arr(1, 2) = "c"
    Arr(1, 3) = "d"
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, False)
    
    'Assert:
    Assert.AreEqual "TestTable", LO.Name
    Assert.AreEqual 8, CInt(LO.Range.Count)
    Assert.AreEqual 4, CInt(LO.Range.Columns.Count)
    Assert.AreEqual 2, CInt(LO.Range.Rows.Count)
    Assert.AreEqual "asdf", LO.Range.Cells(1, 1).Text
    Assert.AreEqual "1234", LO.Range.Cells(1, 2).Text
    Assert.AreEqual "2022/11/02", LO.Range.Cells(1, 3).Text
    Assert.AreEqual "FALSE", LO.Range.Cells(1, 4).Text
    Assert.AreEqual "asdf", LO.ListColumns(1).Name
    Assert.AreEqual "1234", LO.ListColumns(2).Name
    Assert.AreEqual "2022/11/02", LO.ListColumns(3).Name
    Assert.AreEqual "FALSE", LO.ListColumns(4).Name
    
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayToNewTable_1dArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(2) As Variant
    Dim RangeObj As Range
    Dim LO As ListObject
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0) = "col1"
    Arr(1) = "col2"
    Arr(2) = "col3"
    
    Set RangeObj = WB.ActiveSheet.Range("K4")
    Set LO = ArrayToNewTable("TestTable2", Ensure2dArray(Arr), RangeObj, True)
    
    'Assert:
    Assert.AreEqual "TestTable2", LO.Name
    Assert.AreEqual "col2", LO.Range(1, 2).Value
    
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscArray")
Private Sub Test_ArrayToRange_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As New Workbook
    Dim Arr(2) As Variant
    Dim RangeObj As Range
    Dim LO As ListObject
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0) = "col1"
    Arr(1) = "col2"
    Arr(2) = "col3"
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set LO = ArrayToRange(Arr, RangeObj, True)

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

'@TestMethod("MiscArray")
Private Sub Test_ArrayToNewTable_fail()
    Const ExpectedError As Long = -999
    On Error GoTo TestFail
    
    ''Arrange:
    Dim WB As New Workbook
    Dim Arr(1, 2) As Variant
    Dim RangeObj As Range
    Dim LO As ListObject
    
    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "col1"
    Arr(0, 1) = "col2"
    Arr(0, 2) = "col3"
    Arr(1, 0) = "=[d]"
    Arr(1, 1) = "=d"
    Arr(1, 2) = 1
    
    Set RangeObj = WB.ActiveSheet.Range("B4")
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, True)
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, True)

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


'@TestMethod("MiscArray")
Private Sub Test_Ensure2DArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr1() As Variant
    Dim Arr2() As Variant
    
    'Act:
    Arr1 = Array("a", "b", "c")
    Arr1 = Ensure2dArray(Arr1)
    Assert.AreEqual "a", Arr1(0, 0)
    Assert.AreEqual "b", Arr1(0, 1)
    Assert.AreEqual "c", Arr1(0, 2)
    
    ReDim Arr2(0 To 0, 0 To 2)
    Arr2(0, 0) = "a": Arr2(0, 1) = "b": Arr2(0, 2) = "c"
    Arr2 = Ensure2dArray(Arr2)
    Assert.AreEqual "a", Arr2(0, 0)
    Assert.AreEqual "b", Arr2(0, 1)
    Assert.AreEqual "c", Arr2(0, 2)
    
    'Assert:
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_IsArrayAllocated()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ArrTest1() As Variant
    Dim ArrTest2() As Variant
    Dim ArrTest3() As Variant
    Dim ArrTest4(4) As Variant
    Dim ArrTest5() As Variant
    Dim ArrTest6() As Variant
    
    'Act:
    ArrTest1 = Array("a", "b", "c")
    ReDim ArrTest3(10) As Variant
    ArrTest5 = Array("a", "b", "c")
    ReDim ArrTest5(10) As Variant
    ArrTest6 = Array("a", "b", "c")
    ReDim Preserve ArrTest6(10) As Variant
    
    'Assert:
    Assert.IsTrue IsArrayAllocated(ArrTest1)
    Assert.IsFalse IsArrayAllocated(ArrTest2)
    Assert.IsTrue IsArrayAllocated(ArrTest3)
    Assert.IsTrue IsArrayAllocated(ArrTest4)
    Assert.IsTrue IsArrayAllocated(ArrTest5)
    Assert.IsTrue IsArrayAllocated(ArrTest6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_jaggedArrayToLO()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ArrTest(3) As Variant
    Dim WB As Workbook
    Dim LO As ListObject

    'Act:
    Set WB = ExcelBook("", False, False)
    
    ArrTest(0) = Array("col1", "col2", "col3")
    ArrTest(1) = Array(1, 2, 3)
    ArrTest(2) = Array(10)
    ArrTest(3) = Array(100, 200)
    
    Set LO = JaggedArrayToLO(ArrTest, "TableName", WB.Worksheets(1))

    'Assert:
    Assert.AreEqual "TableName", LO.DisplayName
    Assert.AreEqual 2, CInt(LO.DataBodyRange(1, 2).Value)
    Assert.AreEqual 10, CInt(LO.DataBodyRange(2, 1).Value)
    Assert.AreEqual 200, CInt(LO.DataBodyRange(3, 2).Value)

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_jaggedArrayToLO_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim ArrTest(1) As Variant
    Dim WB As Workbook
    Dim LO As ListObject

    'Act:
    Set WB = ExcelBook("", False, False)
    
    ArrTest(0) = Array("col1", "col2", "col3")
    ArrTest(1) = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)

    Set LO = JaggedArrayToLO(ArrTest, "TableName", WB.Worksheets(1))

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

'@TestMethod("MiscArray")
Private Sub Test_IsInArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(3) As Variant
    
    'Act:
    Arr(0) = 1
    Arr(1) = "2"
    Arr(2) = "k"
    Arr(3) = 3.4
    
    'Assert:
    Assert.IsTrue IsInArray(Arr, 1)
    Assert.IsFalse IsInArray(Arr, "1")
    Assert.IsFalse IsInArray(Arr, 2)
    Assert.IsTrue IsInArray(Arr, "2")
    Assert.IsTrue IsInArray(Arr, 3.4)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayDuplicates()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(7) As Variant
    Dim ArrOut() As Variant

    'Act:
    Arr(0) = 1
    Arr(1) = "1"
    Arr(2) = 1
    Arr(3) = 3.4
    Arr(4) = "asdf"
    Arr(5) = 3.4
    Arr(6) = 3.4
    Arr(7) = "1"
    ArrOut = ArrayDuplicates(Arr)
    
    'Assert:
    Assert.AreEqual CLng(4), UBound(ArrOut) - LBound(ArrOut) + 1
    Assert.AreEqual 1, ArrOut(0)
    Assert.AreEqual 3.4, ArrOut(1)
    Assert.AreEqual 3.4, ArrOut(2)
    Assert.AreEqual "1", ArrOut(3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayDuplicates2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(7) As Variant
    Dim ArrOut() As Variant

    'Act:
    Arr(0) = 1
    Arr(1) = "1"
    Arr(2) = 1
    Arr(3) = 3.4
    Arr(4) = "asdf"
    Arr(5) = 3.4
    Arr(6) = 3.4
    Arr(7) = "1"
    ArrOut = ArrayDuplicates(Arr, True)

    'Assert:
    Assert.AreEqual CLng(4), UBound(ArrOut) - LBound(ArrOut) + 1
    Assert.AreEqual "{Entry: 3}: 1", ArrOut(0)
    Assert.AreEqual "{Entry: 6}: 3,4", ArrOut(1)
    Assert.AreEqual "{Entry: 7}: 3,4", ArrOut(2)
    Assert.AreEqual "{Entry: 8}: 1", ArrOut(3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayGetDimension()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr1() As Variant
    Dim Arr2(1) As Variant
    Dim Arr3(1, 1) As Variant
    Dim Arr4(1, 1, 1) As Variant
    Dim Arr5(1, 1, 1, 1) As Variant

    'Assert:
    Assert.AreEqual CLng(0), ArrayGetDimension(Arr1)
    Assert.AreEqual CLng(1), ArrayGetDimension(Arr2)
    Assert.AreEqual CLng(2), ArrayGetDimension(Arr3)
    Assert.AreEqual CLng(3), ArrayGetDimension(Arr4)
    Assert.AreEqual CLng(4), ArrayGetDimension(Arr5)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayUniqueValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(7) As Variant
    Dim ArrOut() As Variant

    'Act:
    Arr(0) = 1
    Arr(1) = "1"
    Arr(2) = 1
    Arr(3) = 3.4
    Arr(4) = "asdf"
    Arr(5) = 3.4
    Arr(6) = 3.4
    Arr(7) = "1"
    ArrOut = ArrayUniqueValues(Arr)
    
    'Assert:
    Assert.AreEqual CLng(4), UBound(ArrOut) - LBound(ArrOut) + 1
    Assert.AreEqual 1, ArrOut(0)
    Assert.AreEqual "1", ArrOut(1)
    Assert.AreEqual 3.4, ArrOut(2)
    Assert.AreEqual "asdf", ArrOut(3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_ArrayUniqueValues_2D()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(1, 2) As Variant
    Dim ArrOut() As Variant

    'Act:
    Arr(0, 0) = 1
    Arr(0, 1) = "1"
    Arr(0, 2) = 1
    Arr(1, 0) = 3.4
    Arr(1, 1) = "asdf"
    Arr(1, 2) = 3.4
    ArrOut = ArrayUniqueValues(Arr)
    
    'Assert:
    Assert.AreEqual CLng(4), UBound(ArrOut) - LBound(ArrOut) + 1
    Assert.AreEqual 1, ArrOut(0)
    Assert.AreEqual "1", ArrOut(1)
    Assert.AreEqual 3.4, ArrOut(2)
    Assert.AreEqual "asdf", ArrOut(3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscArray")
Private Sub Test_isArrayUnique()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr(3) As Variant
    Dim Arr2(3) As Variant

    'Act:
    Arr(0) = 1
    Arr(1) = "1"
    Arr(2) = "3.4"
    Arr(3) = 3.4
    
    Arr2(0) = "asdf"
    Arr2(1) = 3.4
    Arr2(2) = 3.4
    Arr2(3) = "1"
    'Assert:
    Assert.IsTrue IsArrayUnique(Arr)
    Assert.IsFalse IsArrayUnique(Arr2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
