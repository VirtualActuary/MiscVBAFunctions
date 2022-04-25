Attribute VB_Name = "Test__MiscDictionaryCreate"
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

'@TestMethod("MiscDictionaryCreate")
Private Sub Test_dict()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary
    
    'Act:
    Set d = dict("a", 2, "b", ThisWorkbook)

    'Assert:
    Assert.AreEqual 2, d.Item("a")
    Assert.AreEqual ThisWorkbook.Name, d.Item("b").Name

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDictionaryCreate")
Private Sub Test_dicti()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary
    
    'Act:
    Set d = dicti("a", 2, "b", ThisWorkbook)

    'Assert:
    Assert.AreEqual 2, d.Item("a")
    Assert.AreEqual 2, d.Item("A")
    Assert.AreEqual ThisWorkbook.Name, d.Item("b").Name
    Assert.AreEqual ThisWorkbook.Name, d.Item("B").Name

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
