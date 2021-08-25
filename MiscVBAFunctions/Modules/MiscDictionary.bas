Attribute VB_Name = "MiscDictionary"
Option Explicit

Function dict( _
    Optional arg1 As Variant = Empty, Optional arg2 As Variant = Empty, Optional arg3 As Variant = Empty, Optional arg4 As Variant = Empty, _
    Optional arg5 As Variant = Empty, Optional arg6 As Variant = Empty, Optional arg7 As Variant = Empty, Optional arg8 As Variant = Empty, _
    Optional arg9 As Variant = Empty, Optional arg10 As Variant = Empty, Optional arg11 As Variant = Empty, Optional arg12 As Variant = Empty, _
    Optional arg13 As Variant = Empty, Optional arg14 As Variant = Empty, Optional arg15 As Variant = Empty, Optional arg16 As Variant = Empty, _
    Optional arg17 As Variant = Empty, Optional arg18 As Variant = Empty, Optional arg19 As Variant = Empty, Optional arg20 As Variant = Empty, _
    Optional arg21 As Variant = Empty, Optional arg22 As Variant = Empty, Optional arg23 As Variant = Empty, Optional arg24 As Variant = Empty, _
    Optional arg25 As Variant = Empty, Optional arg26 As Variant = Empty, Optional arg27 As Variant = Empty, Optional arg28 As Variant = Empty, _
    Optional arg29 As Variant = Empty, Optional arg30 As Variant = Empty) As Dictionary
    
    Set dict = New Dictionary
    If arg1 = Empty Then Exit Function Else dict.Add arg1, arg2
    If arg3 = Empty Then Exit Function Else dict.Add arg3, arg4
    If arg5 = Empty Then Exit Function Else dict.Add arg5, arg6
    If arg7 = Empty Then Exit Function Else dict.Add arg7, arg8
    If arg9 = Empty Then Exit Function Else dict.Add arg9, arg10
    If arg11 = Empty Then Exit Function Else dict.Add arg11, arg12
    If arg13 = Empty Then Exit Function Else dict.Add arg13, arg14
    If arg15 = Empty Then Exit Function Else dict.Add arg15, arg16
    If arg17 = Empty Then Exit Function Else dict.Add arg17, arg18
    If arg19 = Empty Then Exit Function Else dict.Add arg19, arg20
    If arg21 = Empty Then Exit Function Else dict.Add arg21, arg22
    If arg23 = Empty Then Exit Function Else dict.Add arg23, arg24
    If arg25 = Empty Then Exit Function Else dict.Add arg25, arg26
    If arg27 = Empty Then Exit Function Else dict.Add arg27, arg28
    If arg29 = Empty Then Exit Function Else dict.Add arg29, arg30
End Function

