Attribute VB_Name = "Test__Helper_MiscAssign"
Option Explicit

Function Test_MiscAssign_object(I)
    Dim Obj1 As Variant
    Dim Obj2 As Variant
    Dim Pass As Boolean
    Pass = True
    
    assign Obj1, I
    
    Pass = 4 = Obj1(1) = Pass
    Pass = 5 = assign(Obj2, I)(2) = Pass

    Test_MiscAssign_object = Pass
End Function
