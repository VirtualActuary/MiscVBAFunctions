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

'@TestMethod("MiscDictionary")
Private Sub Test_Concat_Dicts()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Dim d3 As Dictionary

    'Act:
    Set d1 = dict("a", 1, "b", 2)
    Set d2 = dict("c", 10, "d", 20)
    Set d3 = dict(2, 10, "a", 20)
    ConcatDicts d1, d3, d2
    
    'Assert:
    Assert.AreEqual 20, d1("a")
    Assert.AreEqual 2, d1("b")
    Assert.AreEqual 10, d1("c")
    Assert.AreEqual 20, d1("d")
    Assert.AreEqual 10, d1(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDictionary")
Private Sub Test_Join_Dicts()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Dim d3 As Dictionary
    
    'Act:
    Set d1 = dict("a", 1, "b", 2)
    Set d2 = dict("c", 10, "d", 20)
    Set d3 = dict(1, 10, 2, 20)
    Set d = JoinDicts(d1, d2, d3)
    
    'Assert:
    Assert.AreEqual 1, d("a")
    Assert.AreEqual 2, d("b")
    Assert.AreEqual 10, d("c")
    Assert.AreEqual 20, d("d")
    Assert.AreEqual 10, d(1)
    Assert.AreEqual 20, d(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

