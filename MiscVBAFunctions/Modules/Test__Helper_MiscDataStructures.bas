Attribute VB_Name = "Test__Helper_MiscDataStructures"
Option Explicit

Function Test_EnsureUniqueKey_Col()
    Dim Pass As Boolean
    Pass = True
    Dim C1 As Collection
    Dim C2 As Collection
    Set C1 = New Collection
    Set C2 = New Collection

    C1.Add 1, "a"
    C1.Add 1, "b"
    C1.Add 1, "c"
    
    C2.Add 1, "a"
    C2.Add 1, "b"
    C2.Add 1, "b1"
    
    Pass = "d" = EnsureUniqueKey(C1, "d") = Pass = True
    Pass = "b2" = EnsureUniqueKey(C2, "b") = Pass = True
    
    Test_EnsureUniqueKey_Col = Pass
End Function
