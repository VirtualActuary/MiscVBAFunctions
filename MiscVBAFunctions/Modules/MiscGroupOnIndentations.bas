Attribute VB_Name = "MiscGroupOnIndentations"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Sub GroupRowsOnIndentations(R As Range)
    ' groups the rows based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range or Rows that will be grouped.
    
    Dim ri As Range
    For Each ri In R
        ri.EntireRow.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub

Public Sub GroupColumnsOnIndentations(R As Range)
    ' groups the columns based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range of Columns that will be grouped.
    
    Dim ri As Range
    For Each ri In R
        ri.EntireColumn.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub


Private Sub TestRemoveGroupings()
    
    ' Test rows
    RemoveRowGroupings ThisWorkbook.Worksheets("GroupOnIndentations")
    ' Test columns
    RemoveColumnGroupings ThisWorkbook.Worksheets("GroupOnIndentations")
End Sub


Public Sub RemoveRowGroupings(WS As Worksheet)
    ' Remove Row Grouping from the selected Worksheet.
    '
    ' Args:
    '   WS: The workseheet where the grouping will be removed.
    
    Dim R As Range
    Dim ri As Range
    Set R = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels RowLevels:=8
    For Each ri In R.Columns(1)
        ri.EntireRow.OutlineLevel = 1
    Next ri
End Sub

Public Sub RemoveColumnGroupings(WS As Worksheet)
    ' Remove Column Grouping from the selected Worksheet.
    '
    ' Args:
    '   WS: The workseheet where the grouping will be removed.
    
    Dim R As Range
    Dim ri As Range
    Set R = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels columnlevels:=8
    For Each ri In R.Rows(1)
        ri.EntireColumn.OutlineLevel = 1
    Next ri
End Sub

