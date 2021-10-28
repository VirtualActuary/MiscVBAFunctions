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
Private Sub test__HasKey__Collection()                        'TODO Rename test
    On Error GoTo TestFail
    
    
   
    
    'Arrange:
     Dim c As New Collection

    'Act:
    c.Add "foo", "a"
    c.Add col("x", "y", "z"), "b"
    
    'Assert:
    'Assert.AreEqual hasKey(c, "a")
    Assert.AreEqual True, hasKey(c, "a") ' True for scalar
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
