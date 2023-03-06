Attribute VB_Name = "MiscGroupOnIndentations"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Sub GroupRowsOnIndentations(R As Range)
    ' groups the rows based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range or Rows that will be grouped.
    
    Dim Ri As Range
    For Each Ri In R
        Ri.EntireRow.OutlineLevel = Ri.IndentLevel + 1
    Next Ri
    
End Sub

Public Sub GroupColumnsOnIndentations(R As Range)
    ' groups the columns based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range of Columns that will be grouped.
    
    Dim Ri As Range
    For Each Ri In R
        Ri.EntireColumn.OutlineLevel = Ri.IndentLevel + 1
    Next Ri
    
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
    Dim Ri As Range
    Set R = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels RowLevels:=8
    For Each Ri In R.Columns(1)
        Ri.EntireRow.OutlineLevel = 1
    Next Ri
End Sub

Public Sub RemoveColumnGroupings(WS As Worksheet)
    ' Remove Column Grouping from the selected Worksheet.
    '
    ' Args:
    '   WS: The workseheet where the grouping will be removed.
    
    Dim R As Range
    Dim Ri As Range
    Set R = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels Columnlevels:=8
    For Each Ri In R.Rows(1)
        Ri.EntireColumn.OutlineLevel = 1
    Next Ri
End Sub

