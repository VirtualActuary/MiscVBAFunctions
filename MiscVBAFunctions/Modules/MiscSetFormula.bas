Attribute VB_Name = "MiscSetFormula"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Sub SetFormula(ByVal Target As Object, ByVal Formula As String)
    ' Set a formula using `.Formula2` if available. Otherwise use `.Formula`.
    '
    ' Solution from https://stackoverflow.com/a/70849437/6822528
    '
    ' Args:
    '     Target:
    '         The range for which to set the formula. This is typed as `Object` rather than
    '         `Range` to prevent compile errors on older Excel versions.
    '     Formula:
    '         The formula to write to the range.
    
    If Not TypeOf Target Is Range Then Err.Raise 5 ' Type Mismatch: `Target` must be a `Range`.
    
    If HasDynamicArrays Then
        ' This is a late-bound call, which will still compile on older Excel versions.
        Target.Formula2 = Formula
    Else
        Target.Formula = Formula
    End If
End Sub
