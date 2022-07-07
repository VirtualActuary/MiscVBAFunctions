Attribute VB_Name = "Test__MiscOs"
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

'@TestMethod("MiscOs")
Private Sub Test_Path()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual "C:\folder1\folder2\folder3", Path("C:\", "folder1", "folder2", "folder3")
    Assert.AreEqual "C:\folder1\folder2\folder3", Path("C:\", "folder1\", "folder2", "\folder3")
    Assert.AreEqual "C:\folder1\folder2\folder3", Path("C:\", "\folder1\", "\folder2\", "\folder3")
    
    Assert.AreEqual "folder\file.extension", Path("folder", "file.extension")
    Assert.AreEqual "C:\folder1\folder2\folder3", Path("C:\folder1\folder2\folder3")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscOs")
Private Sub Test_is64BitXl()
    On Error GoTo TestFail

    'Assert:
    #If Win64 Then
        Assert.IsTrue Is64BitXl()
    #Else
        Assert.IsFalse Is64BitXl()
    #End If

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscOs")
Private Sub Test_EvalPath()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscOs")
Private Sub Test_CreateFolders()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Fail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
