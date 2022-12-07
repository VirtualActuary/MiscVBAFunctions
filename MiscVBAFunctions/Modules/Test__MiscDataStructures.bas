Attribute VB_Name = "Test__MiscDataStructures"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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

'@TestMethod("MiscDataStructures")
Private Sub Test_EnsureUniqueKey_Col()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    Dim C2 As Collection
    Set C = New Collection
    Set C2 = New Collection

    'Act:
    C.Add 1, "a"
    C.Add 1, "b"
    C.Add 1, "c"
    
    C2.Add 1, "a"
    C2.Add 1, "b"
    C2.Add 1, "b1"
    
    'Assert:
    Assert.AreEqual "d", EnsureUniqueKey(C, "d")
    Assert.AreEqual "b2", EnsureUniqueKey(C2, "b")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDataStructures")
Private Sub Test_EnsureUniqueKey_Dict()
    On Error GoTo TestFail
    
    'Arrange:
    Dim D As Dictionary
    Dim D2 As Dictionary
    
    'Act:
    Set D = Dict("a", 1, "b", 1, "c", 1)
    Set D2 = Dict("a", 1, "b", 1, "b1", 1)
    
    'Assert:
    Assert.AreEqual "d", EnsureUniqueKey(D, "d")
    Assert.AreEqual "b2", EnsureUniqueKey(D2, "b")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
