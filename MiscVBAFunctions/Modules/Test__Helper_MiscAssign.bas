Attribute VB_Name = "Test__Helper_MiscAssign"
Option Explicit

Function Test_MiscAssign_object(I)
    Dim X As Variant
    Dim Y As Variant
    Dim Pass As Boolean
    Pass = True
    
    assign X, I
    
    Pass = 4 = X(1) = Pass
    Pass = 5 = assign(Y, I)(2) = Pass

    Test_MiscAssign_object = Pass
End Function
