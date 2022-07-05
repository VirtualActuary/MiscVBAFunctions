Attribute VB_Name = "Test__MiscListOfChildren"
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

'@TestMethod("MiscListOfChildren")
Private Sub Test_GetListOfChildren()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Depths As Collection

    'Act:
    Set Depths = col( _
    1, _
      2, _
      2, _
        3, _
        3, _
          4, _
      2, _
    1, _
    1, _
      2)
    
    Dim ChildrenDepths As Collection
    ' test down / forward children
    Set ChildrenDepths = GetListOfChildren(Depths)
    
    Assert.AreEqual ChildrenDepths(1)(1), CLng(2)
    Assert.AreEqual ChildrenDepths(1)(2), CLng(3)
    Assert.AreEqual ChildrenDepths(1)(3), CLng(7)
    
    Assert.AreEqual ChildrenDepths(3)(1), CLng(4)
    Assert.AreEqual ChildrenDepths(3)(2), CLng(5)
    
    Assert.AreEqual ChildrenDepths(5)(1), CLng(6)
    
    Assert.AreEqual ChildrenDepths(9)(1), CLng(10)
    
    ' test back / upwards children
    Set ChildrenDepths = GetListOfChildren(Depths, False)
    Assert.AreEqual ChildrenDepths(7)(1), CLng(5)
    Assert.AreEqual ChildrenDepths(7)(2), CLng(4)
    
    Assert.AreEqual ChildrenDepths(8)(1), CLng(7)
    Assert.AreEqual ChildrenDepths(8)(2), CLng(3)
    Assert.AreEqual ChildrenDepths(8)(3), CLng(2)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
