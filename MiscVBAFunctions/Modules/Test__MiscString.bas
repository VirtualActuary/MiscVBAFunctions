Attribute VB_Name = "Test__MiscString"
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

'@TestMethod("MiscString")
Private Sub Test_randomString()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual 4, CInt(Len(randomString(4)))
    Assert.AreNotEqual randomString(5), randomString(5)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscString")
Private Sub Test_EndsWith()
    On Error GoTo TestFail

    'Assert:
    Assert.IsTrue EndsWith("foo bar baz", " baz")
    Assert.IsTrue EndsWith("foo bar baz", "az")
    Assert.IsFalse EndsWith("foo bar baz", " baz ")
    Assert.IsFalse EndsWith("foo bar baz", "bar")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscString")
Private Sub Test_StartsWith()
    On Error GoTo TestFail
    
    'Assert:
    Assert.IsTrue StartsWith("foo bar baz", "foo ")
    Assert.IsTrue StartsWith("foo bar baz", "foo bar baz")
    Assert.IsFalse StartsWith("foo bar baz", "bar")
    Assert.IsFalse StartsWith("foo bar baz", " Foo")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

