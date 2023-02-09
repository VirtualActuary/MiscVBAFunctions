Attribute VB_Name = "Test__MiscHasKey"
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

'@TestMethod("MiscHasKey")
Private Sub Test_HasKey_Collection()
    On Error GoTo TestFail
    
    
    'Arrange:
     Dim C As New Collection

    'Act:
    C.Add "foo", "a"
    C.Add col("x", "y", "z"), "b"
    
    'Assert:
    Assert.AreEqual True, hasKey(C, "a") ' True for scalar
    Assert.AreEqual True, hasKey(C, "b") ' True for scalar
    Assert.AreEqual True, hasKey(C, "A") ' True for case insensitive
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub Test_HasKey_Workbook()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual True, hasKey(Workbooks, ThisWorkbook.Name)
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub Test_HasKey_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim D As New Dictionary
    
    'Act:
    D.Add "a", "foo"
    D.Add "b", col("x", "y", "z")

    'Assert:
    Assert.AreEqual True, hasKey(D, "a") ' True for scalar
    Assert.AreEqual True, hasKey(D, "b") ' True for scalar
    Assert.AreEqual False, hasKey(D, "A") ' False - case sensitive by default
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub Test_HasKey_Dictionary_object()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dObj As Object
    Set dObj = CreateObject("Scripting.Dictionary")
    'Act:
    dObj.Add "a", "foo"
    dObj.Add "b", col("x", "y", "z")

    'Assert:
    Assert.AreEqual True, hasKey(dObj, "a") ' True for scalar
    Assert.AreEqual True, hasKey(dObj, "b") ' True for scalar
    Assert.AreEqual False, hasKey(dObj, "A") ' False - case sensitive by default

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub Test_HasKey_Dictionary_fail()                        'TODO Rename test
    Const ExpectedError As Long = 9              'TODO Change to expected error number
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    hasKey 5, "a"
    hasKey ThisWorkbook, "A"

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
