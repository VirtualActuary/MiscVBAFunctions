Attribute VB_Name = "MiscRowCount"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Function ActiveRowsDown(Optional RangeInput As Range) As Long
    ' number of active rows down from the starting range
    ' it's capped at a minimum of 1
    ' This function takes hidden cells into account
    '
    ' Args:
    '   RangeInput: Starting cell to start counting from
    '
    ' Returns:
    '   The number of rows down from the starting range
    
    Dim WS As Worksheet
    Set WS = RangeInput.Worksheet

    If WS.AutoFilterMode And WS.FilterMode Then
        Dim R_col As Range
        Set R_col = WS.Range(WS.Cells(RangeInput.Row, RangeInput.Column), WS.Cells(WS.Rows.Count, RangeInput.Column))
        
        Dim ArrVals() As Variant
        ReDim ArrVals(R_col.Rows.Count - 1)
        ArrVals = RangeToArray(R_col)
        
        Dim EmptyCellCount As Long
        EmptyCellCount = 0
        
        Dim LastActive As Long
        Dim Counter As Long
        For Counter = LBound(ArrVals) To UBound(ArrVals)
            If IsEmpty(ArrVals(Counter)) Then
                EmptyCellCount = EmptyCellCount + 1
            Else
                LastActive = Counter + 1
                EmptyCellCount = 0
            End If
            
            If EmptyCellCount > 100 Then
                ActiveRowsDown = LastActive
                Exit Function
            End If
        Next Counter
    Else
        If RangeInput Is Nothing Then Set RangeInput = Selection
        ActiveRowsDown = WS.Cells(WS.Cells.Rows.Count, RangeInput.Column).End(XlUp).Row - RangeInput.Row + 1
        ActiveRowsDown = Application.WorksheetFunction.Max(1, ActiveRowsDown)
    End If
    
End Function
