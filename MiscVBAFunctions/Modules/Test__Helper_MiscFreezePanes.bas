Attribute VB_Name = "Test__Helper_MiscFreezePanes"
Option Explicit

Function Test_FreezePanes(RangeObj As Range)
    FreezePanes RangeObj
    
    With Application.Windows(RangeObj.Parent.Parent.Name)
        Test_FreezePanes = .FreezePanes
    End With
End Function


Function Test_UnFreezePanes(RangeObj As Range)
    FreezePanes RangeObj
    UnFreezePanes RangeObj.Parent
    
    With Application.Windows(RangeObj.Parent.Parent.Name)
        Test_UnFreezePanes = Not .FreezePanes
    End With
End Function
