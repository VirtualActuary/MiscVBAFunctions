Attribute VB_Name = "MiscFreezePanes"
'@IgnoreModule ImplicitByRefModifier
Option Explicit


Private Sub Test()
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Worksheets(1)
    FreezePanes WS.Range("D6")
    
    
End Sub

Public Sub FreezePanes(R As Range)
    ' FreezePanes on the current active sheet. Removes FreezedPanes if it already exists.
    '
    ' Args:
    '   r: (row, column) cell where the FreezePanes should occur
    '
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    Dim WS As Worksheet
    Set WS = R.Parent
    
    Dim CurrentScreenUpdating As Boolean
    CurrentScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With Application.Windows(WS.Parent.Name)
        ' if existing freezed panes, remove them
        If .FreezePanes = True Then
            .FreezePanes = False
        End If
        Application.GoTo WS.Cells(1, 1) ' <- to ensure we don't hide the top/ left side of sheet
        ' Unfortunately, we have to do this :/
        Application.GoTo R
        .FreezePanes = True
    End With
    
    Application.ScreenUpdating = CurrentScreenUpdating
    
    CurrentActiveSheet.Activate
End Sub

Public Sub UnFreezePanes(WS As Worksheet)
    '
    '
    ' Args:
    '   WS: Worksheet where this function will execute.
    '

    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    ' Unfortunately, we have to do this :/
    WS.Activate
    With Application.Windows(WS.Parent.Name)
        .FreezePanes = False
    End With
    
    CurrentActiveSheet.Activate
End Sub
