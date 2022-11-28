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
Private Sub Test_ExpandEnvironmentalVariables()
    On Error GoTo TestFail
   
    'Assert:
    Assert.AreEqual Environ("windir"), ExpandEnvironmentalVariables("%windir%")
    Assert.AreEqual Environ("username"), ExpandEnvironmentalVariables("%username%")
    Assert.AreEqual Environ("windir") & "\%foo\bar%\%username\" & Environ("username"), ExpandEnvironmentalVariables("%windir%\%foo\bar%\%username\%username%")

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
    Dim Dir As String
    
    'Act:
    Dir = Path(ExpandEnvironmentalVariables("%temp%"), "folder1", "folder2", "folder3")
    MakeDirs Dir
    
    'Assert:
    Assert.IsTrue fso.FolderExists(Dir)
    
TestExit:
    fso.DeleteFolder Path(ExpandEnvironmentalVariables("%temp%"), "folder1")
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscOs")
Private Sub Test_RunShell()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual 0, CInt(RunShell("cmd /c echo hello", True))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscOs")
Private Sub Test_CreateFolders_fail()
    Const ExpectedError As Long = 53
    On Error GoTo TestFail

    'Act:
    MakeDirs ("\\MadeUpDir\Foo")

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

'@TestMethod("MiscOs")
Private Sub Test_CreateFolders_fail_2()
    Const ExpectedError As Long = 58
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dir As String
    
    'Act:
    Dir = Path(ExpandEnvironmentalVariables("%temp%"), "folder1", "folder2", "folder3")
    MakeDirs Dir
    MakeDirs Dir, False

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
