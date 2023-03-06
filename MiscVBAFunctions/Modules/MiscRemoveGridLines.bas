Attribute VB_Name = "MiscRemoveGridLines"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Sub RemoveGridLines(WS As Worksheet)
    ' Remove all GridLines from the selected Worksheet.
    '
    ' Args:
    '   WS: Selected WorkSheet.
    
    Dim View As WorksheetView
    For Each View In WS.Parent.Windows(1).SheetViews
        If View.Sheet.Name = WS.Name Then
            View.DisplayGridlines = False
            Exit Sub
        End If
    Next
End Sub
