Attribute VB_Name = "Test__MiscFSO"
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

'@TestMethod("MiscFSO")
Private Sub Test_GetAllFilesRecursive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AllFiles As Collection
    
    'Act:
    Set AllFiles = GetAllFilesRecursive(fso.GetFolder(Path(ThisWorkbook.Path, "\tests\GetAllFiles")))

    'Assert:
    Assert.AreEqual 5, CInt(AllFiles.Count)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
