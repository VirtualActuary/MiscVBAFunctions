Attribute VB_Name = "Test__MiscRemoveGridlines"
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

'@TestMethod("MisRemoveGridlines")
Private Sub TestRemoveGridlines()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    RemoveGridLines ThisWorkbook.Sheets(1)
    
    'Act:

    'Assert:
    Dim GridlinesShown As Boolean
    GridlinesShown = ThisWorkbook.Windows(1).SheetViews(1).DisplayGridlines
    Assert.AreEqual False, GridlinesShown

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
