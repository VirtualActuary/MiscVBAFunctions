Attribute VB_Name = "Test__Helper_MiscNewKeys"
Option Explicit

Function Test_GetNewKey1()
    Dim Pass As Boolean
    Pass = True
    
    Dim C As New Collection
    Dim I As Long

    C.Add "bla", "name"
    For I = 1 To 20
        C.Add "bla", "name" & I
    Next I

    Pass = "name21" = GetNewKey("name", C) = Pass = True
    Pass = "NewName" = GetNewKey("NewName", C) = Pass = True
    
    Test_GetNewKey1 = Pass
End Function


Function Test_GetNewKey2()
    Dim Pass As Boolean
    Pass = True

    Dim D As New Collection

    D.Add "bla", "does"
    D.Add "bla", "not"
    D.Add "bla", "matter"

    Pass = "not1" = GetNewKey("not", D) = Pass = True
    Pass = "foo" = GetNewKey("foo", D) = Pass = True
    
    Test_GetNewKey2 = Pass
End Function
