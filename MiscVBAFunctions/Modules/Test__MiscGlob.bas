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
    Dim c As Collection
    
    
    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "folder1")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(1))
    
    
    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "folder[1-9]")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(2))
    
    
    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "folder?")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(2))
    
  
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
    Dim c As Collection
    
    
    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "*")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\empty file.txt", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(3))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", "*")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(1))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "*\*er[1-9]\*")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(2))

    
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
    Dim c As Collection
    
    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "folder1\**")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(2))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "**\*er[1-9]\*.xlsx")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(1))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "*\*er[1-9]\**")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(3))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "**\*")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\empty file.txt", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\empty file.txt", CStr(c(3))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(4))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(5))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(6))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\empty file.docx", CStr(c(7))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1", CStr(c(8))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(9))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1\empty file.txt", CStr(c(10))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "**")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(3))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(4))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1", CStr(c(5))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(6))


    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "*\folder1\**\*.xlsx")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(1))

    
    
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
    Dim c As Collection

    'Act:
    Set c = Glob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "**\**")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(3))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2", CStr(c(4))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1", CStr(c(5))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(6))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod2()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As Collection
    
    'Act:
    Set c = RGlob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "folder1")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1", CStr(c(1))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1", CStr(c(2))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1", CStr(c(3))
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder2\folder1\folder1", CStr(c(4))
    
    'Act:
    Set c = RGlob("C:\AA\MiscVBAFunctions\tests\GetAllFiles", "*.xlsx")
    'Assert:
    Assert.AreEqual "C:\AA\MiscVBAFunctions\tests\GetAllFiles\folder1\folder1\empty file.xlsx", CStr(c(1))


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod3()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod4()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
