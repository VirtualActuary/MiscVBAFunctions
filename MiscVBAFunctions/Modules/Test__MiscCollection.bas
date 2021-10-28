Attribute VB_Name = "Test__MiscCollection"
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

'@TestMethod("MiscCollection.min")
Private Sub Test_min()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 4, min(col(4, 5, 6)), "min test succeeded"
    Assert.AreEqual 5, min(col(5, 6)), "min test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.min")
Private Sub TestMethod1()                        'TODO Rename test
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As Collection
    'Act:
    
    
    min c
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Assert.Succeed
        
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
