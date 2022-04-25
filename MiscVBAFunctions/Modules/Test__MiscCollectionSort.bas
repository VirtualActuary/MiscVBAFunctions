Attribute VB_Name = "Test__MiscCollectionSort"
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

'@TestMethod("MiscCollectionSort")
Private Sub Test_BubbleSort()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coll As Collection
    
    'Act:
    Set coll = col("variables10", "variables", "variables2", "variables_10", "variables_2")
    Set coll = BubbleSort(coll)

    'Assert:
    Assert.AreEqual coll(1), "variables"
    Assert.AreEqual coll(2), "variables10"
    Assert.AreEqual coll(3), "variables2"
    Assert.AreEqual coll(4), "variables_10"
    Assert.AreEqual coll(5), "variables_2"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
