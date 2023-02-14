Attribute VB_Name = "Test__MiscTextFile"
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

'@TestMethod("MiscTestFile")
Private Sub Test_CreateTextFile()
    On Error GoTo TestFail
    
    'Arrange:
    Dim iFile As Integer
    iFile = FreeFile
    Dim FilePath As String
    FilePath = ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt"
    Dim inputText As String
    inputText = "my test text."
    Dim textline As String
    
    'Act:
    CreateTextFile inputText, FilePath
    
    'Assert:
    
    Open FilePath For Input As #iFile
        Line Input #iFile, textline
        Assert.AreEqual inputText, textline
    Close #iFile

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscTestFile")
Private Sub Test_readTextFile()
    On Error GoTo TestFail
    
    
    'Arrange:
    Dim text As String
    
    'Act:
    text = readTextFile(ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt")
    
    'Assert:
    Assert.AreEqual "my test text." & vbNewLine, text
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
