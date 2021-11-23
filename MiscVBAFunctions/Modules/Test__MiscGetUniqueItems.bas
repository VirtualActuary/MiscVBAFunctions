Attribute VB_Name = "Test__MiscGetUniqueItems"
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

'@TestMethod("MiscGetUniqueItems")
Private Sub Test_GetUniqueItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim arr1(3) As Variant
    Dim arr2(3) As Variant
    Dim arr3(3) As Variant
    Dim arr4(3) As Variant
    Dim arr5(3) As Variant
    
    'Act:
    arr1(0) = "a": arr1(1) = "b": arr1(2) = "c": arr1(3) = "b"
    arr2(0) = "a": arr2(1) = "b": arr2(2) = "c": arr2(3) = "B"
    arr3(0) = "a": arr3(1) = "b": arr3(2) = "c": arr3(3) = "B"
    arr4(0) = 1: arr4(1) = 2: arr4(2) = 3: arr4(3) = 2
    arr5(0) = 1: arr5(1) = 1: arr5(2) = "a": arr5(3) = "a"
    
    'Assert:
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(arr1))  ' zero index
    Assert.AreEqual CLng(3), UBound(GetUniqueItems(arr2), 1) ' zero index + case sensitive
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(arr3, False), 1) ' zero index + case insensitive
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(arr4), 1) ' zero index
    Assert.AreEqual CLng(1), UBound(GetUniqueItems(arr5), 1) ' zero index

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
