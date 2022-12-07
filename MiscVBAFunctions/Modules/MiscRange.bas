Attribute VB_Name = "MiscRange"
Option Explicit

Function ActiveRowsDown(Optional R As Range) As Long
    ' number of active rows down from the starting range
    ' it's capped at a minimum of 1
    '
    ' Args:
    '   R: Input range
    '
    ' Returns:
    '   Number of rows down from the starting range.
    
    If R Is Nothing Then Set R = Selection
    ActiveRowsDown = R.Worksheet.Cells(R.Worksheet.Cells.Rows.Count, R.Column).End(xlUp).Row - R.Row + 1
    ActiveRowsDown = Application.WorksheetFunction.max(1, ActiveRowsDown)
End Function


Sub InsertColumns(R As Range, Optional NrCols As Integer = 1, Optional ShiftToLeft As Boolean = True)
    ' Insert 1 or more column into a range.
    '
    ' Args:
    '   R: Input Range. This Range will be altered.
    '   NrCols: Number of columns to add
    '   ShiftToLeft: Optional -
    '
    ' Returns:
    '
    
    Dim I As Integer
    For I = 1 To NrCols
        If ShiftToLeft Then
            R.EntireColumn.Insert xlShiftToLeft
        Else
            R.EntireColumn.Insert xlShiftToRight
        End If
    Next I
End Sub
