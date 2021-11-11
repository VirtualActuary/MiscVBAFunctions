Attribute VB_Name = "Test__MiscDictionary"
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

'@TestMethod("MiscDictionary")
Private Sub Test_dictget()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary
    

    'On Error Resume Next
    'Debug.Print dictget(d, "c")
    'Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    'On Error GoTo 0

    'Act:
    Set d = dict("a", 2, "b", ThisWorkbook)

    'Assert:
    Assert.AreEqual 2, dictget(d, "a")
    Assert.AreEqual ThisWorkbook.Name, dictget(d, "b").Name
    Assert.AreEqual vbNullString, dictget(d, "c", vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDictionary")
Private Sub Test_dictget_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary

    'Act:
    Set d = dict("a", 2, "b", ThisWorkbook)

    dictget d, "c"
    
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
