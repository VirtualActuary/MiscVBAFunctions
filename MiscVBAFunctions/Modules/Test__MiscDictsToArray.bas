Attribute VB_Name = "Test__MiscDictsToArray"
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

'@TestMethod("MiscDictsToArray")
Private Sub Test_DictsToArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C1 As Collection
    Dim D1 As Dictionary
    Dim D2 As Dictionary
    Dim Arr() As Variant
    Dim ColCounter As Long
    Dim DictCounter As Long
    
    'Act:
    Set D1 = Dict("a", 1, "b", 2, "c", 3)
    Set D2 = Dict("a", 11, "b", 22, "c", 33)
    Set C1 = col(D1, D2)
    Arr = DictsToArray(C1)

    'Assert:
    Assert.AreEqual "a", Arr(0, 0)
    Assert.AreEqual "b", Arr(0, 1)
    Assert.AreEqual "c", Arr(0, 2)
    Assert.AreEqual 1, CInt(Arr(1, 0))
    Assert.AreEqual 2, CInt(Arr(1, 1))
    Assert.AreEqual 3, CInt(Arr(1, 2))
    Assert.AreEqual 11, CInt(Arr(2, 0))
    Assert.AreEqual 22, CInt(Arr(2, 1))
    Assert.AreEqual 33, CInt(Arr(2, 2))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
