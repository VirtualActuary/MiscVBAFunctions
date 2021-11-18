Attribute VB_Name = "MiscRemoveGridLines"
Option Explicit

Public Sub RemoveGridLines(WS As Worksheet)
    Dim view As WorksheetView
    For Each view In WS.Parent.Windows(1).SheetViews
        If view.Sheet.Name = WS.Name Then
            view.DisplayGridlines = False
            Exit Sub
        End If
    Next
End Sub
