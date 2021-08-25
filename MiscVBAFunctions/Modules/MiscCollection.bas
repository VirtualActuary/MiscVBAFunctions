Attribute VB_Name = "MiscCollection"
Option Explicit

Function collect( _
    Optional arg1 As Variant = Empty, Optional arg2 As Variant = Empty, Optional arg3 As Variant = Empty, Optional arg4 As Variant = Empty, _
    Optional arg5 As Variant = Empty, Optional arg6 As Variant = Empty, Optional arg7 As Variant = Empty, Optional arg8 As Variant = Empty, _
    Optional arg9 As Variant = Empty, Optional arg10 As Variant = Empty, Optional arg11 As Variant = Empty, Optional arg12 As Variant = Empty, _
    Optional arg13 As Variant = Empty, Optional arg14 As Variant = Empty, Optional arg15 As Variant = Empty, Optional arg16 As Variant = Empty, _
    Optional arg17 As Variant = Empty, Optional arg18 As Variant = Empty, Optional arg19 As Variant = Empty, Optional arg20 As Variant = Empty, _
    Optional arg21 As Variant = Empty, Optional arg22 As Variant = Empty, Optional arg23 As Variant = Empty, Optional arg24 As Variant = Empty, _
    Optional arg25 As Variant = Empty, Optional arg26 As Variant = Empty, Optional arg27 As Variant = Empty, Optional arg28 As Variant = Empty, _
    Optional arg29 As Variant = Empty)
    
    Set collect = New Collection
    If arg1 = Empty Then Exit Function Else collect.Add arg1
    If arg2 = Empty Then Exit Function Else collect.Add arg2
    If arg3 = Empty Then Exit Function Else collect.Add arg3
    If arg4 = Empty Then Exit Function Else collect.Add arg4
    If arg5 = Empty Then Exit Function Else collect.Add arg5
    If arg6 = Empty Then Exit Function Else collect.Add arg6
    If arg7 = Empty Then Exit Function Else collect.Add arg7
    If arg8 = Empty Then Exit Function Else collect.Add arg8
    If arg9 = Empty Then Exit Function Else collect.Add arg9
    If arg10 = Empty Then Exit Function Else collect.Add arg10
    If arg11 = Empty Then Exit Function Else collect.Add arg11
    If arg12 = Empty Then Exit Function Else collect.Add arg12
    If arg13 = Empty Then Exit Function Else collect.Add arg13
    If arg14 = Empty Then Exit Function Else collect.Add arg14
    If arg15 = Empty Then Exit Function Else collect.Add arg15
    If arg16 = Empty Then Exit Function Else collect.Add arg16
    If arg17 = Empty Then Exit Function Else collect.Add arg17
    If arg18 = Empty Then Exit Function Else collect.Add arg18
    If arg19 = Empty Then Exit Function Else collect.Add arg19
    If arg20 = Empty Then Exit Function Else collect.Add arg20
    If arg21 = Empty Then Exit Function Else collect.Add arg21
    If arg22 = Empty Then Exit Function Else collect.Add arg22
    If arg23 = Empty Then Exit Function Else collect.Add arg23
    If arg24 = Empty Then Exit Function Else collect.Add arg24
    If arg25 = Empty Then Exit Function Else collect.Add arg25
    If arg26 = Empty Then Exit Function Else collect.Add arg26
    If arg27 = Empty Then Exit Function Else collect.Add arg27
    If arg28 = Empty Then Exit Function Else collect.Add arg28
    If arg29 = Empty Then Exit Function Else collect.Add arg29
End Function

