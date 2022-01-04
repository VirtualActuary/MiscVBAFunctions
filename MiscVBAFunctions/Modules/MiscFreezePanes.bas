Attribute VB_Name = "MiscFreezePanes"
'@IgnoreModule ImplicitByRefModifier
Option Explicit


Private Sub test()
    
    
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Worksheets(1)
    FreezePanes WS.Range("D6")
    
    
End Sub

Public Sub FreezePanes(r As Range)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    Dim WS As Worksheet
    Set WS = r.Parent
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With Application.Windows(WS.Parent.Name)
        ' if existing freezed panes, remove them
        If .FreezePanes = True Then
            .FreezePanes = False
        End If
        Application.Goto WS.Cells(1, 1) ' <- to ensure we don't hide the top/ left side of sheet
        ' Unfortunately, we have to do this :/
        Application.Goto r
        .FreezePanes = True
    End With
    
    Application.ScreenUpdating = currentScreenUpdating
    
    CurrentActiveSheet.Activate
End Sub

Public Sub UnFreezePanes(WS As Worksheet)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    ' Unfortunately, we have to do this :/
    WS.Activate
    With Application.Windows(WS.Parent.Name)
        .FreezePanes = False
    End With
    
    CurrentActiveSheet.Activate
End Sub
