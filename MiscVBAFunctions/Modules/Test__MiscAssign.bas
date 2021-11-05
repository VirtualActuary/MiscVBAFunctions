Attribute VB_Name = "Test__MiscAssign"
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

'@TestMethod("MiscAssign")
Private Sub Test_MiscAssign_variant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim i As Integer
    

    'Act:

    'Assert:
    Assert.AreEqual 5, assign(i, 5), "assign test succeeded"
    Assert.AreEqual 1.4, assign(i, 1.4), "assign test succeeded"
    
    
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscAssign")
Private Sub Test_MiscAssign_object()
    On Error GoTo TestFail
    
    'Arrange:
    Dim x As Variant
    Dim y As Variant
    Dim i As Variant
    Set i = col(4, 5, 6)
    assign x, i
    
    'Assert:
    Assert.AreEqual 4, x(1)
    Assert.AreEqual 5, assign(y, i)(2)


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
