Attribute VB_Name = "MiscHasDynamicArrays"
Option Explicit


Public Function HasDynamicArrays() As Boolean
    ' Checks whether the Excel/VBA version has dynamic arrays functionality
    ' This will mean it has the Range.Formula2 property which supports dynamic arrays
    '
    ' Returns:
    '    True, if the active Excel has dynamic arrays, otherwise False
    
    ' Solution from https://stackoverflow.com/a/70849437/6822528
    Static IsDynamic As Boolean
    Static RanCheck As Boolean
    
    If Not RanCheck Then
        IsDynamic = Not IsError(Evaluate("=COUNT(@{1,2,3})"))
        RanCheck = True
    End If
    HasDynamicArrays = IsDynamic
End Function
