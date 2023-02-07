Attribute VB_Name = "Test__Helper_MiscAssign"
Option Explicit

Function Test_MiscAssign_object(I)
    Dim x As Variant
    Dim y As Variant
    Dim Pass As Boolean
    Pass = True
    
    assign x, I
    
    Pass = 4 = x(1) = Pass
    Pass = 5 = assign(y, I)(2) = Pass

    Test_MiscAssign_object = Pass
End Function
