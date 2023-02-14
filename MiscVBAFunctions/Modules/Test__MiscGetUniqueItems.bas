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
    Dim Arr1(3) As Variant
    Dim Arr2(3) As Variant
    Dim Arr3(3) As Variant
    Dim Arr4(3) As Variant
    Dim Arr5(3) As Variant
    
    'Act:
    Arr1(0) = "a": Arr1(1) = "b": Arr1(2) = "c": Arr1(3) = "b"
    Arr2(0) = "a": Arr2(1) = "b": Arr2(2) = "c": Arr2(3) = "B"
    Arr3(0) = "a": Arr3(1) = "b": Arr3(2) = "c": Arr3(3) = "B"
    Arr4(0) = 1: Arr4(1) = 2: Arr4(2) = 3: Arr4(3) = 2
    Arr5(0) = 1: Arr5(1) = 1: Arr5(2) = "a": Arr5(3) = "a"
    
    'Assert:
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(Arr1))  ' zero index
    Assert.AreEqual CLng(3), UBound(GetUniqueItems(Arr2), 1) ' zero index + case sensitive
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(Arr3, False), 1) ' zero index + case insensitive
    Assert.AreEqual CLng(2), UBound(GetUniqueItems(Arr4), 1) ' zero index
    Assert.AreEqual CLng(1), UBound(GetUniqueItems(Arr5), 1) ' zero index

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
