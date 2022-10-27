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
    
    Dim arr2(3)
    
    'Act:
    Arr(0, 0) = 100.2: Arr(0, 1) = 1.9
    Arr(1, 0) = 2.1: Arr(1, 1) = 2.2
    EnsureDotSeparatorTransformation Arr

    arr2(0) = 1.2: arr2(1) = 2.1: arr2(2) = 3.8
    EnsureDotSeparatorTransformation arr2
    
    
    'Assert:
    Assert.AreEqual "100.2", Arr(0, 0)
    Assert.AreEqual "1.9", Arr(0, 1)
    Assert.AreEqual "2.1", Arr(1, 0)
    Assert.AreEqual "2.2", Arr(1, 1)
    
    Assert.AreEqual "1.2", arr2(0)
    Assert.AreEqual "2.1", arr2(1)
    Assert.AreEqual "3.8", arr2(2)
    
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
    Dim col1 As Collection
    'Act:
    Arr(0) = 10
    Arr(1) = 11
    Arr(2) = 12
    Arr(3) = 13
    Set col1 = ArrayToCollection(Arr)
    'Assert:
    Assert.AreEqual 10, CInt(col1(1))
    Assert.AreEqual 11, CInt(col1(2))
    Assert.AreEqual 12, CInt(col1(3))
    Assert.AreEqual 13, CInt(col1(4))

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
    Dim arr1() As Variant
    Dim arr2() As Variant
    
    'Act:
    arr1 = Array("a", "b", "c")
    arr1 = Ensure2dArray(arr1)
    Assert.AreEqual "a", arr1(0, 0)
    Assert.AreEqual "b", arr1(0, 1)
    Assert.AreEqual "c", arr1(0, 2)
    
    ReDim arr2(0 To 0, 0 To 2)
    arr2(0, 0) = "a": arr2(0, 1) = "b": arr2(0, 2) = "c"
    arr2 = Ensure2dArray(arr2)
    Assert.AreEqual "a", arr2(0, 0)
    Assert.AreEqual "b", arr2(0, 1)
    Assert.AreEqual "c", arr2(0, 2)
    
    'Assert:
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
