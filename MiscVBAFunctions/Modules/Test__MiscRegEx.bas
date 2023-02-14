Attribute VB_Name = "Test__MiscRegEx"
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

'@TestMethod("MiscRegEx")
Private Sub Test_RenameVariableInFormula()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual "xyz + a ^ xyz + foo(xyz) + abc + abc1 /xyz", RenameVariableInFormula("Ab + a ^ ab + foo(AB) + abc + abc1 /aB", "ab", "xyz")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscRegEx")
Private Sub Test_RenameVariableInFormula_2()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual "FOO + bar - foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "foo", "FOO")

    Assert.AreEqual "foo + bar - foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "fo", "FOO")
    
    Assert.AreEqual "foo + bar - FOOBAR " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "foobar", "FOOBAR")
    
    Assert.AreEqual "FOO + bar - foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "FOO", "FOO")

    Assert.AreEqual "foo + BAR - foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "bar", "BAR")
    
    Assert.AreEqual "foo+BAR-foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo+bar-foobar " & """foo bar foobar" & """", "bar", "BAR")

    Assert.AreEqual "FOO + bar - foobar " & """FOO bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "foo", "FOO", False)

    Assert.AreEqual "foo + bar - foobar " & """foo bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "fo", "FOO", False)

    Assert.AreEqual "foo + bar - FOOBAR " & """foo bar FOOBAR" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "foobar", "FOOBAR", False)

    Assert.AreEqual "FOO + bar - foobar " & """FOO bar foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "FOO", "FOO", False)

    Assert.AreEqual "foo + BAR - foobar " & """foo BAR foobar" & """", RenameVariableInFormula("foo + bar - foobar " & """foo bar foobar" & """", "bar", "BAR", False)
    
    Assert.AreEqual "foo+BAR-foobar " & """foo BAR foobar" & """", RenameVariableInFormula("foo+bar-foobar " & """foo bar foobar" & """", "bar", "BAR", False)


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
