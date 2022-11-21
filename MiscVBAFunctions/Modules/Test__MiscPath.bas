Attribute VB_Name = "Test__MiscPath"
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


'@TestMethod("MiscPath.Path")
Private Sub Test_Path()
    On Error GoTo TestFail
    
    ' Empty path should return empty path
    Assert.AreEqual vbNullString, Path("")
    Assert.AreEqual vbNullString, Path(Array(""))
    Assert.AreEqual vbNullString, Path(col(""))
    
    ' Normalize slashes to '\'.
    Assert.AreEqual "C:\foo\bar", Path("C:", "foo", "bar")
    Assert.AreEqual "C:\foo\bar", Path("C:/", "foo", "bar")
    Assert.AreEqual "C:\foo\bar", Path("C:\", "foo", "bar")
    Assert.AreEqual "C:\foo\bar", Path("C:", "foo/bar/")
    Assert.AreEqual "C:\foo\bar", Path("C:", "foo\bar\")
    Assert.AreEqual "foo\bar\asdf\zxcv", Path("foo/bar", "asdf\zxcv")
    Assert.AreEqual "foo\bar\asdf\zxcv", Path("foo//bar", "asdf\\zxcv")
    
    ' Path segments starting with drive letters throw away what comes before them.
    Assert.AreEqual "D:\bar", Path("C:", "foo", "D:", "bar")
    Assert.AreEqual "D:", Path("C:", "foo", "D:")
    Assert.AreEqual "E:", Path("C:", "D:", "E:")
    
    ' Path segments starting with slashes throw away what comes before them, but preserve the drive letter.
    Assert.AreEqual "C:\bar", Path("C:", "foo", "/bar")
    Assert.AreEqual "C:", Path("C:", "foo", "/")
    
    ' File extensions are preserved
    Assert.AreEqual "folder\file.extension", Path("folder", "file.extension")
    
    ' Single path segment is allowed
    Assert.AreEqual "C:\folder1\folder2\folder3", Path("C:/folder1/folder2/folder3")
    
    ' Collection is allowed in stead of multiple arguments
    Assert.AreEqual "foo\bar\baz", Path(col("foo", "bar", "baz"))
    
    ' Array is allowed in stead of multiple arguments
    Assert.AreEqual "foo\bar\baz", Path(Array("foo", "bar", "baz"))
    
    ' Network paths
    Assert.AreEqual "\\server1\foo", Path("\\server1", "foo")
    
    ' Root paths without drive letters
    Assert.AreEqual "\foo\bar", Path("/foo", "bar")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.IsAbsolutePath")
Private Sub Test_IsAbsolutePath()
    On Error GoTo TestFail
    
    'Assert:
    Assert.IsTrue IsAbsolutePath("C:")
    Assert.IsTrue IsAbsolutePath("C:\")
    Assert.IsTrue IsAbsolutePath("C:/")
    Assert.IsTrue IsAbsolutePath("\asdf")
    Assert.IsTrue IsAbsolutePath("/asdf")
    Assert.IsTrue IsAbsolutePath("\")
    Assert.IsTrue IsAbsolutePath("/")
    
    Assert.IsFalse IsAbsolutePath("asdf")
    Assert.IsFalse IsAbsolutePath("asdf/")
    Assert.IsFalse IsAbsolutePath("asdf\")
    Assert.IsFalse IsAbsolutePath("asdf/foo")
    Assert.IsFalse IsAbsolutePath("asdf\foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.PathGetDrive")
Private Sub Test_PathGetDrive()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual "C:", PathGetDrive("C:")
    Assert.AreEqual "C:", PathGetDrive("C:\")
    Assert.AreEqual "C:", PathGetDrive("C:/")
    Assert.AreEqual "C:", PathGetDrive("C:\asdf")
    Assert.AreEqual "C:", PathGetDrive("C:/asdf")
    Assert.AreEqual "", PathGetDrive("\asdf")
    Assert.AreEqual "", PathGetDrive("/asdf")
    Assert.AreEqual "", PathGetDrive("\")
    Assert.AreEqual "", PathGetDrive("/")
    Assert.AreEqual "", PathGetDrive("asdf")
    Assert.AreEqual "", PathGetDrive("asdf/")
    Assert.AreEqual "", PathGetDrive("asdf\")
    Assert.AreEqual "", PathGetDrive("asdf/foo")
    Assert.AreEqual "", PathGetDrive("asdf\foo")
    Assert.AreEqual "", PathGetDrive("\\server1\foo")
    Assert.AreEqual "", PathGetDrive("//server1/foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.PathHasDrive")
Private Sub Test_PathHasDrive()
    On Error GoTo TestFail
    
    'Assert:
    Assert.IsTrue PathHasDrive("C:")
    Assert.IsTrue PathHasDrive("C:\")
    Assert.IsTrue PathHasDrive("C:/")
    Assert.IsTrue PathHasDrive("C:\asdf")
    Assert.IsTrue PathHasDrive("C:/asdf")
    Assert.IsFalse PathHasDrive("\asdf")
    Assert.IsFalse PathHasDrive("/asdf")
    Assert.IsFalse PathHasDrive("\")
    Assert.IsFalse PathHasDrive("/")
    Assert.IsFalse PathHasDrive("asdf")
    Assert.IsFalse PathHasDrive("asdf/")
    Assert.IsFalse PathHasDrive("asdf\")
    Assert.IsFalse PathHasDrive("asdf/foo")
    Assert.IsFalse PathHasDrive("asdf\foo")
    Assert.IsFalse PathHasDrive("\\server1\foo")
    Assert.IsFalse PathHasDrive("//server1/foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.PathGetServer")
Private Sub Test_PathGetServer()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual "", PathGetServer("C:")
    Assert.AreEqual "", PathGetServer("C:\")
    Assert.AreEqual "", PathGetServer("C:/")
    Assert.AreEqual "", PathGetServer("C:\asdf")
    Assert.AreEqual "", PathGetServer("C:/asdf")
    Assert.AreEqual "", PathGetServer("\asdf")
    Assert.AreEqual "", PathGetServer("/asdf")
    Assert.AreEqual "", PathGetServer("\")
    Assert.AreEqual "", PathGetServer("/")
    Assert.AreEqual "", PathGetServer("asdf")
    Assert.AreEqual "", PathGetServer("asdf/")
    Assert.AreEqual "", PathGetServer("asdf\")
    Assert.AreEqual "", PathGetServer("asdf/foo")
    Assert.AreEqual "", PathGetServer("asdf\foo")
    Assert.AreEqual "\\server1", PathGetServer("\\server1\foo")
    Assert.AreEqual "//server1", PathGetServer("//server1/foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.PathHasServer")
Private Sub Test_PathHasServer()
    On Error GoTo TestFail
    
    'Assert:
    Assert.IsFalse PathHasServer("C:")
    Assert.IsFalse PathHasServer("C:\")
    Assert.IsFalse PathHasServer("C:/")
    Assert.IsFalse PathHasServer("C:\asdf")
    Assert.IsFalse PathHasServer("C:/asdf")
    Assert.IsFalse PathHasServer("\asdf")
    Assert.IsFalse PathHasServer("/asdf")
    Assert.IsFalse PathHasServer("\")
    Assert.IsFalse PathHasServer("/")
    Assert.IsFalse PathHasServer("asdf")
    Assert.IsFalse PathHasServer("asdf/")
    Assert.IsFalse PathHasServer("asdf\")
    Assert.IsFalse PathHasServer("asdf/foo")
    Assert.IsFalse PathHasServer("asdf\foo")
    Assert.IsTrue PathHasServer("\\server1\foo")
    Assert.IsTrue PathHasServer("//server1/foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscPath.AbsolutePath")
Private Sub Test_AbsolutePath()
    On Error GoTo TestFail

    'Assert:
    'Absolute paths
    Assert.AreEqual "C:\", AbsolutePath("C:\")
    Assert.AreEqual "C:\test", AbsolutePath("C:\test")
    Assert.AreEqual "C:\foo\bar", AbsolutePath("C:\foo\bar")
    Assert.AreEqual "C:\bar", AbsolutePath("C:\foo\..\bar")
    Assert.AreEqual "C:\foo", AbsolutePath("C:\foo\\bar\\..")
    Assert.AreEqual PathGetDrive(ThisWorkbook.Path) & "\foo", AbsolutePath("\foo")
    
    ' Relative paths
    Assert.AreEqual ThisWorkbook.Path & "\foo\bar", AbsolutePath("foo/bar", ThisWorkbook)
    Assert.AreEqual ThisWorkbook.Path & "\foo", AbsolutePath(".\foo", ThisWorkbook)
    Assert.AreEqual fso.GetParentFolderName(ThisWorkbook.Path) & "\foo", AbsolutePath("..\foo", ThisWorkbook)
    Assert.AreEqual fso.GetParentFolderName(ThisWorkbook.Path), AbsolutePath("foo\..\..", ThisWorkbook)
    Assert.AreEqual fso.GetParentFolderName(ThisWorkbook.Path), AbsolutePath("foo//..//..", ThisWorkbook)
    
    ' Network Paths
    Assert.AreEqual "\\foo\bar", AbsolutePath("\\foo/bar")
    Assert.AreEqual "\\foo\bar", AbsolutePath("\\foo\\bar")
    Assert.AreEqual "\\foo\bar", AbsolutePath("\\foo\\.\bar\\..\\bar")
    Assert.AreEqual "\\foo", AbsolutePath("//foo")
    Assert.AreEqual "\\foo\bar", AbsolutePath("//foo\\bar")
    Assert.AreEqual "\\hello\2", AbsolutePath("\\hello\world\\..\2")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscPath.EvalPath")
Private Sub Test_EvalPath()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual "C:\foo", EvalPath("C:\foo")
    Assert.AreEqual "C:\foo", EvalPath("C:/foo")
    Assert.AreEqual "C:\c", EvalPath("C:\a\..\b\..\c")
    Assert.AreEqual Environ("HOMEDRIVE") & "\Users\" & Environ("username"), EvalPath("%HOMEDRIVE%\Users\%username%")
    Assert.AreEqual Path(ThisWorkbook.Path, "foo\bar"), EvalPath("foo/bar")
    Assert.AreEqual Path(ThisWorkbook.Path, "foo\" & Environ("username")), EvalPath("foo/%UserName%")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
