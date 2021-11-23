Attribute VB_Name = "Test__MiscArray"
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

'@TestMethod("Uncategorized")
Private Sub TestMethod1()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim tst As Variant
    Dim I As Long, J As Long
    Dim arr(2, 2)
    
    Dim arr2(3)
    
    'Act:
    arr(0, 0) = 100.2: arr(0, 1) = 1.9
    arr(1, 0) = 2.1: arr(1, 1) = 2.2
    EnsureDotSeparatorTransformation arr

    arr2(0) = 1.2: arr2(1) = 2.1: arr2(2) = 3.8
    EnsureDotSeparatorTransformation arr2
    
    'Assert:
    Assert.AreEqual "100.2", arr(0, 0)
    Assert.AreEqual "1.9", arr(0, 1)
    Assert.AreEqual "2.1", arr(1, 0)
    Assert.AreEqual "2.2", arr(1, 1)
    
    Assert.AreEqual "1.2", arr2(0)
    Assert.AreEqual "2.1", arr2(1)
    Assert.AreEqual "3.8", arr2(2)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
