Attribute VB_Name = "Test__MiscDate"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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

'@TestMethod("MiscDate")
Private Sub Test_EoWeek()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual "2022 December 04", Format(EoWeek(CDate("December 1, 2022"), 0), "yyyy mmmm dd")
    Assert.AreEqual "2022 December 11", Format(EoWeek(CDate("December 1, 2022"), 1), "yyyy mmmm dd")
    Assert.AreEqual "2022 November 27", Format(EoWeek(CDate("December 1, 2022"), -1), "yyyy mmmm dd")
    
    ' Different formats
    Assert.AreEqual "2020 January 12", Format(EoWeek("2020/01/10", 0), "yyyy mmmm dd")
    Assert.AreEqual "2020 January 12", Format(EoWeek("1/10/2020", 0), "yyyy mmmm dd")  ' USA format
    Assert.AreEqual "2010 May 02", Format(EoWeek(40279, 3), "yyyy mmmm dd")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
