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
    Dim d As Dictionary
    Set d = New Dictionary
    
    Dim dI As Dictionary
    Set dI = New Dictionary
    
    'Act:
    d.Add "a", 1  ' lowercase
    Set dI = EnsureDictI(d)

    'Assert:
    Assert.AreEqual False, d.Exists("A")
    Assert.AreEqual True, d.Exists("a")
    
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
    Dim c As Collection
    Set c = New Collection
    
    Dim cI As Collection
    Set cI = New Collection
    
    Dim d1 As Dictionary
    Set d1 = New Dictionary
    
    Dim d2 As Dictionary
    Set d2 = New Dictionary
    
    Dim d3 As Dictionary
    Set d3 = New Dictionary
    
    'Act:
    d1.Add "A", "foo"  ' uppercase
    d2.Add "b", "foo"  ' lowercase
    d3.Add "C", "foo"  ' uppercase
    
    c.Add d1
    c.Add d2
    c.Add d3
    
    Set cI = EnsureDictI(c)

    'Assert:
    Assert.AreEqual False, c(1).Exists("a")
    Assert.AreEqual True, c(2).Exists("b")
    Assert.AreEqual False, c(3).Exists("c")
    
    
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
