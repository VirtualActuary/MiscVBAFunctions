Attribute VB_Name = "Test__Helper_MiscGroupOnIndent"
Option Explicit

Function TestGroupOnIndentationsRows(WB)
    Dim Pass As Boolean
    Pass = True
    
    Dim RowR As Range
    Set RowR = WB.Names("__TestGroupRowsOnIndentations__").RefersToRange
    
    GroupRowsOnIndentations RowR

    Pass = CLng(1) = CLng(RowR(1).EntireRow.OutlineLevel) = Pass
    Pass = CLng(2) = CLng(RowR(2).EntireRow.OutlineLevel) = Pass
    Pass = CLng(2) = CLng(RowR(3).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(4).EntireRow.OutlineLevel) = Pass
    Pass = CLng(2) = CLng(RowR(5).EntireRow.OutlineLevel) = Pass
    Pass = CLng(3) = CLng(RowR(6).EntireRow.OutlineLevel) = Pass
    Pass = CLng(3) = CLng(RowR(7).EntireRow.OutlineLevel) = Pass

    TestGroupOnIndentationsRows = Pass
End Function


Function TestGroupOnIndentationsColumns(WB)
    Dim Pass As Boolean
    Pass = True

    Dim ColR As Range
    Set ColR = WB.Names("__TestGroupColumnsOnIndentations__").RefersToRange

    GroupColumnsOnIndentations ColR

    Pass = CLng(1) = CLng(ColR(1).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(2) = CLng(ColR(2).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(3) = CLng(ColR(3).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(ColR(4).EntireColumn.OutlineLevel) = Pass
    
    TestGroupOnIndentationsColumns = Pass
End Function


Function TestUnGroupOnIndentationsRow(WB)
    Dim Pass As Boolean
    Pass = True
    
    Dim RowR As Range
    Set RowR = WB.Names("__TestGroupRowsOnIndentations__").RefersToRange
    
    RemoveRowGroupings WB.Sheets(1)

    Pass = CLng(1) = CLng(RowR(1).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(2).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(3).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(4).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(5).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(6).EntireRow.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(RowR(7).EntireRow.OutlineLevel) = Pass
   
   TestUnGroupOnIndentationsRow = Pass
End Function


Function TestUnGroupOnIndentationsCol(WB)
    Dim Pass As Boolean
    Pass = True
  
    Dim ColR As Range
    Set ColR = WB.Names("__TestGroupColumnsOnIndentations__").RefersToRange
    
    RemoveColumnGroupings WB.Sheets(1)
    
    Pass = CLng(1) = CLng(ColR(1).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(ColR(2).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(ColR(3).EntireColumn.OutlineLevel) = Pass
    Pass = CLng(1) = CLng(ColR(4).EntireColumn.OutlineLevel) = Pass
    
    TestUnGroupOnIndentationsCol = Pass
End Function




