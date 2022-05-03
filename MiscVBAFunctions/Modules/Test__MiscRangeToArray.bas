Attribute VB_Name = "Test__MiscRangeToArray"
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
    Set WB = Workbooks.Open(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscRangeToArray\RangeToArray.xlsx"), ReadOnly:=False)
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

'@TestMethod("MiscRangeToArray")
Private Sub Test_RangeToArray_2D()
    On Error GoTo TestFail
    
    'Arrange:
    'Dim WB As Workbook
    Dim WS As Worksheet
    Dim x As Range
    Dim y() As Variant
    
    'Act:
    Set WS = WB.Sheets(1)
    Set x = WS.Range("A1:C2")
    y = RangeToArray(x)

    'Assert:
    Assert.AreEqual 11, CInt(y(1, 1))
    Assert.AreEqual 12, CInt(y(1, 2))
    Assert.AreEqual 9, CInt(y(1, 3))
    Assert.AreEqual 13, CInt(y(2, 1))
    Assert.AreEqual 14, CInt(y(2, 2))
    Assert.AreEqual 15, CInt(y(2, 3))
   
    Assert.AreEqual 2, CInt(UBound(y) - LBound(y) + 1)
    Assert.AreEqual 3, CInt(UBound(y, 2) - LBound(y, 2) + 1)
    

TestExit:
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscRangeToArray")
Private Sub Test_RangeToArray_1_value()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WS As Worksheet
    Dim x As Range
    Dim y() As Variant
    
    'Act:
    Set WS = WB.Sheets(2)
    Set x = WS.Range("A1")
    y = RangeToArray(x)


    'Assert:
    Assert.AreEqual 4, CInt(y(0))
    Assert.AreEqual 1, CInt(UBound(y) - LBound(y) + 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscRangeToArray")
Private Sub Test_RangeToArray_1D_row()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WS As Worksheet
    Dim x As Range
    Dim y() As Variant
    
    'Act:
    Set WS = WB.Sheets(3)
    Set x = WS.Range("A1:C1")
    y = RangeToArray(x)

    'Assert:
    Assert.AreEqual 1, CInt(y(0))
    Assert.AreEqual 2, CInt(y(1))
    Assert.AreEqual 3, CInt(y(2))
    Assert.AreEqual 3, CInt(UBound(y) - LBound(y) + 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscRangeToArray")
Private Sub Test_RangeToArray_1D_column()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WS As Worksheet
    Dim x As Range
    Dim y() As Variant
    
    'Act:
    Set WS = WB.Sheets(4)
    Set x = WS.Range("A1:A3")
    y = RangeToArray(x)

    'Assert:
    Assert.AreEqual 66, CInt(y(0))
    Assert.AreEqual 77, CInt(y(1))
    Assert.AreEqual 88, CInt(y(2))
    Assert.AreEqual 3, CInt(UBound(y) - LBound(y) + 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

