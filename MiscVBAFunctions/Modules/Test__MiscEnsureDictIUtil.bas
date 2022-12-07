Attribute VB_Name = "Test__MiscEnsureDictIUtil"
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

'@TestMethod("MiscEnsureDictIUtil")
Private Sub Test_EnsureDictI()
    On Error GoTo TestFail
    
    'Arrange:
    Dim D As Dictionary
    Set D = New Dictionary
    
    Dim dI As Dictionary
    Set dI = New Dictionary
    
    'Act:
    D.Add "a", 1  ' lowercase
    Set dI = EnsureDictI(D)

    'Assert:
    Assert.AreEqual False, D.Exists("A")
    Assert.AreEqual True, D.Exists("a")
    
    Assert.AreEqual True, dI.Exists("A")
    Assert.AreEqual True, dI.Exists("a")
    
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscEnsureDictIUtil")
Private Sub Test_EnsureDictIContainer()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    Set C = New Collection
    
    Dim cI As Collection
    Set cI = New Collection
    
    Dim D1 As Dictionary
    Set D1 = New Dictionary
    
    Dim D2 As Dictionary
    Set D2 = New Dictionary
    
    Dim d3 As Dictionary
    Set d3 = New Dictionary
    
    'Act:
    D1.Add "A", "foo"  ' uppercase
    D2.Add "b", "foo"  ' lowercase
    d3.Add "C", "foo"  ' uppercase
    
    C.Add D1
    C.Add D2
    C.Add d3
    
    Set cI = EnsureDictI(C)

    'Assert:
    Assert.AreEqual False, C(1).Exists("a")
    Assert.AreEqual True, C(2).Exists("b")
    Assert.AreEqual False, C(3).Exists("c")
    
    
    Assert.AreEqual True, cI(1).Exists("a")
    Assert.AreEqual True, cI(2).Exists("b")
    Assert.AreEqual True, cI(3).Exists("c")
    
    Assert.AreEqual True, cI(1).Exists("A")
    Assert.AreEqual True, cI(2).Exists("B")
    Assert.AreEqual True, cI(3).Exists("C")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
