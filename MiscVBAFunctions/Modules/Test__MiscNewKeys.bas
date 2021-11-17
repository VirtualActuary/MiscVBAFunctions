Attribute VB_Name = "Test__MiscNewKeys"
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

'@TestMethod("MiscNewKeys")
Private Sub Test_GetNewKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New Collection
    Dim d As New Collection
    Dim I As Long

    'Act:
    c.Add "bla", "name"
    For I = 1 To 100
        c.Add "bla", "name" & I
    Next I
    
    d.Add "bla", "does"
    d.Add "bla", "not"
    d.Add "bla", "matter"

    'Assert:
    Assert.AreEqual "name101", GetNewKey("name", c)
    Assert.AreEqual "NewName", GetNewKey("NewName", c)
    Assert.AreEqual "not1", GetNewKey("not", d)
    Assert.AreEqual "foo", GetNewKey("foo", d)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
