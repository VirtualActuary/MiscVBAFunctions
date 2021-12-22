Attribute VB_Name = "Test__MiscGroupOnIndentations"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private WB As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, "tests\MiscGroupOnIndentations\MiscGroupOnIndentations.xlsx"), True, True)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    WB.Close False
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("MiscGroupOnIndentations")
Private Sub TestGroupOnIndentations()
    On Error GoTo TestFail
    
    'Arrange:
    Dim RowR As Range
    Set RowR = WB.Names("__TestGroupRowsOnIndentations__").RefersToRange
    
    Dim ColR As Range
    Set ColR = WB.Names("__TestGroupColumnsOnIndentations__").RefersToRange
    
    'Act:
    
    ' test rows
    GroupRowsOnIndentations RowR
    ' test columns
    GroupColumnsOnIndentations ColR
    
    'Assert:
    Assert.AreEqual CLng(1), CLng(RowR(1).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(2), CLng(RowR(2).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(2), CLng(RowR(3).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(4).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(2), CLng(RowR(5).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(3), CLng(RowR(6).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(3), CLng(RowR(7).EntireRow.OutlineLevel)
    
    Assert.AreEqual CLng(1), CLng(ColR(1).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(2), CLng(ColR(2).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(3), CLng(ColR(3).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(ColR(4).EntireColumn.OutlineLevel)
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MiscGroupOnIndentations")
Private Sub TestUnGroupOnIndentations()
    On Error GoTo TestFail
    
    'Arrange:
    Dim RowR As Range
    Set RowR = WB.Names("__TestGroupRowsOnIndentations__").RefersToRange
    
    Dim ColR As Range
    Set ColR = WB.Names("__TestGroupColumnsOnIndentations__").RefersToRange
    
    'Act:
    
    ' Test rows
    RemoveRowGroupings WB.Sheets(1)
    ' Test columns
    RemoveColumnGroupings WB.Sheets(1)
    
    'Assert:
    Assert.AreEqual CLng(1), CLng(RowR(1).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(2).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(3).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(4).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(5).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(6).EntireRow.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(RowR(7).EntireRow.OutlineLevel)
    
    Assert.AreEqual CLng(1), CLng(ColR(1).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(ColR(2).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(ColR(3).EntireColumn.OutlineLevel)
    Assert.AreEqual CLng(1), CLng(ColR(4).EntireColumn.OutlineLevel)
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
