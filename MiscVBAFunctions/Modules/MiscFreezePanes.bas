Attribute VB_Name = "MiscFreezePanes"
Option Explicit


Private Sub test()
    
    On Error GoTo UnFreeze
    
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Sheets("Sheet1")
    FreezePanes WS.Range("D4")
    
UnFreeze:
    UnFreezePanes WS
    
End Sub

Sub FreezePanes(r As Range)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    Dim WS As Worksheet
    Set WS = r.Parent
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With Application.Windows(WS.Parent.Name)
        ' Unfortunately, we have to do this :/
        Application.Goto r
        .FreezePanes = True
    End With
    
    Application.ScreenUpdating = currentScreenUpdating
    
    CurrentActiveSheet.Activate
End Sub

Sub UnFreezePanes(WS As Worksheet)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    ' Unfortunately, we have to do this :/
    WS.Activate
    With Application.Windows(WS.Parent.Name)
        .FreezePanes = False
    End With
    
    CurrentActiveSheet.Activate
End Sub