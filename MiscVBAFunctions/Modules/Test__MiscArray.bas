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
    Dim arr(2, 2)
    
    Dim arr2(3)
    
    'Act:
    arr(0, 0) = 100.2: arr(0, 1) = 1.9
    arr(1, 0) = 2.1: arr(1, 1) = 2.2
    EnsureDotSeparatorTransformation arr

    arr2(0) = 1.2: arr2(1) = 2.1: arr2(2) = 3.8
    EnsureDotSeparatorTransformation arr2
    
    
    'Assert:
    Assert.AreEqual "100.2", arr(0, 0)
    Assert.AreEqual "1.9", arr(0, 1)
    Assert.AreEqual "2.1", arr(1, 0)
    Assert.AreEqual "2.2", arr(1, 1)
    
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
    Dim arr(2, 2)
    Dim arrSecond(3)
    
    'Act:
    arr(0, 0) = 100.2: arr(0, 1) = CVErr(xlErrName)
    arr(1, 0) = 2.1: arr(1, 1) = CVErr(xlErrNA)
    ErrorToNullStringTransformation arr

    arrSecond(0) = 1.2: arrSecond(1) = CVErr(xlErrRef): arrSecond(2) = 3.8
    ErrorToNullStringTransformation arrSecond


    'Assert:
    Assert.AreEqual 100.2, arr(0, 0)
    Assert.AreEqual vbNullString, arr(0, 1)
    Assert.AreEqual 2.1, arr(1, 0)
    Assert.AreEqual vbNullString, arr(1, 1)
    
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
    Dim arr(2, 2)
    Dim arrSecond(3)
    Dim arrThird(1)
    Dim arrFourth(1)
    Dim arrFifth(1)

    'Act:
    arr(0, 0) = CDate("2021-1-2"): arr(0, 1) = CDate("2021-01-28 10:2")
    arr(1, 0) = 13: arr(1, 1) = 2.5
    DateToStringTransformation arr

    arrSecond(0) = 1.2: arrSecond(1) = 2.1: arrSecond(2) = CDate("2021-3-28 10:2:10")
    DateToStringTransformation arrSecond
    
    arrThird(0) = CDate("2021-01-28 10:2:10")
    arrFourth(0) = CDate("2021-01-28 10:2:10")
    arrFifth(0) = CDate("2021-01-28 10:2:10")
    
    'Assert:
    Assert.AreEqual "2021-01-02", arr(0, 0)
    Assert.AreEqual "2021-01-28", arr(0, 1)
    Assert.AreEqual 13, arr(1, 0)
    Assert.AreEqual 2.5, arr(1, 1)
    
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
