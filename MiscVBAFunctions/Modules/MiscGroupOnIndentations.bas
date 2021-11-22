Attribute VB_Name = "MiscGroupOnIndentations"
Option Explicit

Private Sub TestGroupOnIndentations()

    ' test rows
    GroupRowsOnIndentations ThisWorkbook.Names("__TestGroupRowsOnIndentations__").RefersToRange
    ' test columns
    GroupColumnsOnIndentations ThisWorkbook.Names("__TestGroupColumnsOnIndentations__").RefersToRange

End Sub

Public Sub GroupRowsOnIndentations(r As Range)
    ' groups the rows based on indentations of the cells in the range
    
    ' Dim ri As Range, WS As Worksheet
    Dim ri As Range
    For Each ri In r
        ri.EntireRow.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub


Public Sub GroupColumnsOnIndentations(r As Range)
    ' groups the columns based on indentations of the cells in the range
    
    ' Dim ri As Range, WS As Worksheet
    Dim ri As Range
    For Each ri In r
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
    Dim r As Range
    Dim ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    For Each ri In r.Columns(1)
        ri.EntireRow.OutlineLevel = 1
    Next ri
End Sub

Public Sub RemoveColumnGroupings(WS As Worksheet)
    Dim r As Range
    Dim ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    For Each ri In r.Rows(1)
        ri.EntireColumn.OutlineLevel = 1
    Next ri
End Sub

