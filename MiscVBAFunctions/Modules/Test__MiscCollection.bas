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
Private Sub Test_min()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 4, min(col(7, 4, 5, 6)), "min test succeeded"
    Assert.AreEqual 5, min(col(9, 5, 6)), "min test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.min")
Private Sub Test_min_fail()
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

'@TestMethod("MiscCollection.max")
Private Sub Test_max()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 6, max(col(4, 5, 6, 1, 2)), "max test succeeded"
    Assert.AreEqual 6.1, max(col(5.3, 6.1)), "max test succeeded"


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.max")
Private Sub Test_max_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As Collection

    'Act:
    max c

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection.mean")
Private Sub Test_mean()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 4#, mean(col(4, 5, 6, 3, 2)), "mean test succeeded"
    Assert.AreEqual 6#, mean(col(5, 7)), "mean test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.mean")
Private Sub Test_mean_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As Collection

    'Act:
    mean c

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
