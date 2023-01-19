Attribute VB_Name = "Test__MiscCsv"
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

'@TestMethod("MiscCsv")
Private Sub Test_CsvToLO()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LO As ListObject
    Dim WB As Workbook
    Dim WS As Worksheet

    'Act:
    Set WB = ExcelBook("")
    Set WS = WB.Worksheets(1)
    Set LO = CsvToLO(WS.Cells(10, 10), Fso.BuildPath(ThisWorkbook.Path, ".\tests\Csv\MiscCsv.csv"), "MyTable")
    
    'Assert:
    Assert.AreEqual "MyTable", LO.Name
    Assert.AreEqual "1", LO.Range(1, 1).Value
   
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

