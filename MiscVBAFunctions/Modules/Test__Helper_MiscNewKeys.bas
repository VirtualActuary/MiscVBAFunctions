Attribute VB_Name = "Test__Helper_MiscNewKeys"
Option Explicit

Function Test_GetNewKey()
    Dim Pass As Boolean
    Pass = True
    
    Dim C As New Collection
    Dim d As New Collection
    Dim I As Long

    'Act:
    C.Add "bla", "name"
    For I = 1 To 100
        C.Add "bla", "name" & I
    Next I
    
    d.Add "bla", "does"
    d.Add "bla", "not"
    d.Add "bla", "matter"

    'Assert:
    Pass = "name101" = GetNewKey("name", C) = Pass
    Pass = "NewName" = GetNewKey("NewName", C) = Pass
    Pass = "not1" = GetNewKey("not", d) = Pass
    Pass = "foo" = GetNewKey("foo", d) = Pass
    
    Test_GetNewKey = Pass
End Function
