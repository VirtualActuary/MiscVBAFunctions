Attribute VB_Name = "Test__MiscErrorMessage"
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

'@TestMethod("MiscErrorMessage")
Private Sub Test_ErrorMessage()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    
    'Assert:
    Assert.AreEqual "This array is fixed or temporarily locked", ErrorMessage(10)
    Assert.AreEqual "Out of memory: a fix is required before continuing", ErrorMessage(7, "a fix is required before continuing")
    Assert.AreEqual "Unknown error", ErrorMessage(77)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
