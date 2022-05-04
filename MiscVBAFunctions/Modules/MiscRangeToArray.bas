Attribute VB_Name = "MiscRangeToArray"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function RangeToArray(r As Range, _
                Optional IgnoreEmptyInFlatArray As Boolean) As Variant()
    ' Converts a range to a normalized array.
    ' vectors allocated to 1-dimensional arrays
    ' tables allocated to 2-dimensional array
    '
    ' Args:
    '   r: Range to be converted to an array.
    '   IgnoreEmptyInFlatArray: If True, skip over empty results.
    '
    ' Returns:
    '   The normalized array.
    
    If r.Cells.Count = 1 Then
        RangeToArray = Array(r.Value)
    ElseIf r.Rows.Count = 1 Or r.Columns.Count = 1 Then
        RangeToArray = RangeTo1DArray(r, IgnoreEmptyInFlatArray)
    Else
        RangeToArray = RangeTo2DArray(r)
    End If
End Function

Public Function RangeTo1DArray( _
              r As Range _
            , Optional IgnoreEmpty As Boolean = True _
            ) As Variant()
    ' currently does the same as rangeToArray, just named better and is more efficient
    ' instead of reading from memory for every range item, we read it in only once
    '
    ' Args:
    '   r: Range to be converted to an array.
    '   IgnoreEmpty: If True, skip over empty results.
    '
    ' Returns:
    '   The normalized array.
    
    Dim arr() As Variant ' the output array
    ReDim arr(r.Cells.Count - 1)
    
    Dim Values() As Variant ' values of the whole range
    If r.Cells.Count = 1 Then
        arr(0) = r.Value
        RangeTo1DArray = arr
        Exit Function
    End If
    
    Values = r.Value
    Dim I As Long
    Dim J As Long
    Dim counter As Long
    counter = 0
    For I = LBound(Values, 1) To UBound(Values, 1) ' rows
        For J = LBound(Values, 2) To UBound(Values, 2) ' columns
            If IsError(Values(I, J)) Then
                ' if error, we cannot check if empty, we need to add it
                arr(counter) = Values(I, J)
                counter = counter + 1
            ElseIf Values(I, J) = vbNullString And IgnoreEmpty Then
                ReDim Preserve arr(UBound(arr) - 1) ' when there is an empty cell, just reduce array size by 1
            Else
                arr(counter) = Values(I, J)
                counter = counter + 1
            End If
        Next J
    Next I
    
    RangeTo1DArray = arr
    
End Function


Public Function RangeTo2DArray(r As Range) As Variant()
    ' ensure a range is converted to a 2-dimensional array
    ' special treatment on edge cases where a range is a 1x1 scalar
    '
    ' Args:
    '   r: Range to be converted to an array.
    '
    ' Returns:
    '   2D array.
    
    If r.Cells.Count = 1 Then
        Dim arr_single() As Variant
        ReDim arr_single(1 To 1, 1 To 1) ' make it base 1, similar to what .value does for non-scalars
        arr_single(1, 1) = r.Value
        RangeTo2DArray = arr_single
        Exit Function
    End If
    
    Dim Values() As Variant ' values of the whole range
    Values = r.Value

    Dim arr() As Variant ' the output array
    ReDim arr(UBound(Values, 1) - LBound(Values, 1), UBound(Values, 2) - LBound(Values, 2))
    Dim I As Long
    Dim J As Long
    Dim I_start As Long
    Dim J_start As Long
    I_start = LBound(Values, 1)
    J_start = LBound(Values, 2)
    For I = LBound(Values, 1) To UBound(Values, 1) ' rows
        For J = LBound(Values, 2) To UBound(Values, 2) ' columns
            arr(I - I_start, J - J_start) = Values(I, J)
        Next J
    Next I
    RangeTo2DArray = arr
    
End Function
