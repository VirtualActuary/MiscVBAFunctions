Attribute VB_Name = "Test__MiscCollectionCreate"
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

'@TestMethod("MiscCollectionCreate")
Private Sub Test_Col()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    'Act:
    Set C = Col(1, 3, 5)
    'Assert:
    'Assert.Succeed
    
    Assert.AreEqual 1, C(1), "col test succeeded"
    Assert.AreEqual 3, C(2), "col test succeeded"
    Assert.AreEqual 5, C(3), "col test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollectionCreate")
Private Sub Test_zip()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C1 As Collection
    Dim C2 As Collection
    Dim Cout As Collection

    'Act:
    Set C1 = Col(1, 2, 3)
    Set C2 = Col(4, 5, 6, 7)
    
    Set Cout = Zip(C1, C2)

    'Assert:
    Assert.AreEqual 1, Cout(1)(1), "zip test succeeded"
    Assert.AreEqual 4, Cout(1)(2), "zip test succeeded"
    
    Assert.AreEqual 2, Cout(2)(1), "zip test succeeded"
    Assert.AreEqual 5, Cout(2)(2), "zip test succeeded"
    
    Assert.AreEqual 3, Cout(3)(1), "zip test succeeded"
    Assert.AreEqual 6, Cout(3)(2), "zip test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
