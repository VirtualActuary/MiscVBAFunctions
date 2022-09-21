Attribute VB_Name = "Test__MiscGlob"
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

'@TestMethod("MiscGlob")
Private Sub Test_Glob_1()  ' Simple tests
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder1")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder[1-9]")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(2))
    
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder?")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(2))
    
  
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscGlob")
Private Sub Test_Glob_2()  ' "*" tests
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\empty file.txt", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(3))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", "*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "*\*er[1-9]\*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(2))

    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscGlob")
Private Sub Test_Glob_3()  ' 1x recursion in the pattern
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder1\**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**\*er[1-9]\*.xlsx")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "*\*er[1-9]\**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(3))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**\*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\empty file.txt", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\empty file.txt", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(6))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\empty file.docx", CStr(C(7))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(8))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(9))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(10))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(6))


    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "*\folder1\**\*.xlsx")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))

    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("MiscGlob")
Private Sub Test_Glob_4()  ' Multiple recursion in the pattern
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection

    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**\**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(6))
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**\folder1\**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(4))
    
    'Act:
    Set C = Glob(ThisWorkbook.Path & "\tests\GetAllFiles", "**\folder1\**\*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\empty file.txt", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(5))
     

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscGlob")
Private Sub Test_RGlob_1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder1")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(4))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "*.xlsx")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "*.txt")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\empty file.txt", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\empty file.txt", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(3))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "*1*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(4))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder[1-9]")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(5))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder?")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(5))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscGlob")
Private Sub Test_RGlob_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles\folder2", "*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\empty file.docx", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(3))
     Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(4))
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(6))

    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(6))


    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\empty file.txt", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\empty file.txt", CStr(C(3))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1", CStr(C(4))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(5))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(6))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\empty file.docx", CStr(C(7))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(8))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(9))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(10))

    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "*\*er[1-9]\*")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(C(3))


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscGlob")
Private Sub Test_RGlob_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    
    
    'Act:
    Set C = RGlob(ThisWorkbook.Path & "\tests\GetAllFiles", "folder2\**")
    'Assert:
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2", CStr(C(1))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1", CStr(C(2))
    Assert.AreEqual ThisWorkbook.Path & "\tests\GetAllFiles\folder2\folder1\folder1", CStr(C(3))
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

