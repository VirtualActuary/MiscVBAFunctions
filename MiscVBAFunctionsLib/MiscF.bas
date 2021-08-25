Attribute VB_Name = "MiscF"
Option Explicit

'************"MiscCollection"


Function collect( _
    Optional arg1 As Variant = Empty, Optional arg2 As Variant = Empty, Optional arg3 As Variant = Empty, Optional arg4 As Variant = Empty, _
    Optional arg5 As Variant = Empty, Optional arg6 As Variant = Empty, Optional arg7 As Variant = Empty, Optional arg8 As Variant = Empty, _
    Optional arg9 As Variant = Empty, Optional arg10 As Variant = Empty, Optional arg11 As Variant = Empty, Optional arg12 As Variant = Empty, _
    Optional arg13 As Variant = Empty, Optional arg14 As Variant = Empty, Optional arg15 As Variant = Empty, Optional arg16 As Variant = Empty, _
    Optional arg17 As Variant = Empty, Optional arg18 As Variant = Empty, Optional arg19 As Variant = Empty, Optional arg20 As Variant = Empty, _
    Optional arg21 As Variant = Empty, Optional arg22 As Variant = Empty, Optional arg23 As Variant = Empty, Optional arg24 As Variant = Empty, _
    Optional arg25 As Variant = Empty, Optional arg26 As Variant = Empty, Optional arg27 As Variant = Empty, Optional arg28 As Variant = Empty, _
    Optional arg29 As Variant = Empty) As Collection
    
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


Function zip( _
    Optional arg1 As Collection = Nothing, Optional arg2 As Collection = Nothing, Optional arg3 As Collection = Nothing, Optional arg4 As Collection = Nothing, _
    Optional arg5 As Collection = Nothing, Optional arg6 As Collection = Nothing, Optional arg7 As Collection = Nothing, Optional arg8 As Collection = Nothing, _
    Optional arg9 As Collection = Nothing, Optional arg10 As Collection = Nothing, Optional arg11 As Collection = Nothing, Optional arg12 As Collection = Nothing, _
    Optional arg13 As Collection = Nothing, Optional arg14 As Collection = Nothing, Optional arg15 As Collection = Nothing, Optional arg16 As Collection = Nothing, _
    Optional arg17 As Collection = Nothing, Optional arg18 As Collection = Nothing, Optional arg19 As Collection = Nothing, Optional arg20 As Collection = Nothing, _
    Optional arg21 As Collection = Nothing, Optional arg22 As Collection = Nothing, Optional arg23 As Collection = Nothing, Optional arg24 As Collection = Nothing, _
    Optional arg25 As Collection = Nothing, Optional arg26 As Collection = Nothing, Optional arg27 As Collection = Nothing, Optional arg28 As Collection = Nothing, _
    Optional arg29 As Collection = Nothing) As Collection
    

    Dim AllArgs As Collection
    Set AllArgs = New Collection
    If arg1 Is Nothing Then GoTo AllDone Else AllArgs.Add arg1
    If arg2 Is Nothing Then GoTo AllDone Else AllArgs.Add arg2
    If arg3 Is Nothing Then GoTo AllDone Else AllArgs.Add arg3
    If arg4 Is Nothing Then GoTo AllDone Else AllArgs.Add arg4
    If arg5 Is Nothing Then GoTo AllDone Else AllArgs.Add arg5
    If arg6 Is Nothing Then GoTo AllDone Else AllArgs.Add arg6
    If arg7 Is Nothing Then GoTo AllDone Else AllArgs.Add arg7
    If arg8 Is Nothing Then GoTo AllDone Else AllArgs.Add arg8
    If arg9 Is Nothing Then GoTo AllDone Else AllArgs.Add arg9
    If arg10 Is Nothing Then GoTo AllDone Else AllArgs.Add arg10
    If arg11 Is Nothing Then GoTo AllDone Else AllArgs.Add arg11
    If arg12 Is Nothing Then GoTo AllDone Else AllArgs.Add arg12
    If arg13 Is Nothing Then GoTo AllDone Else AllArgs.Add arg13
    If arg14 Is Nothing Then GoTo AllDone Else AllArgs.Add arg14
    If arg15 Is Nothing Then GoTo AllDone Else AllArgs.Add arg15
    If arg16 Is Nothing Then GoTo AllDone Else AllArgs.Add arg16
    If arg17 Is Nothing Then GoTo AllDone Else AllArgs.Add arg17
    If arg18 Is Nothing Then GoTo AllDone Else AllArgs.Add arg18
    If arg19 Is Nothing Then GoTo AllDone Else AllArgs.Add arg19
    If arg20 Is Nothing Then GoTo AllDone Else AllArgs.Add arg20
    If arg21 Is Nothing Then GoTo AllDone Else AllArgs.Add arg21
    If arg22 Is Nothing Then GoTo AllDone Else AllArgs.Add arg22
    If arg23 Is Nothing Then GoTo AllDone Else AllArgs.Add arg23
    If arg24 Is Nothing Then GoTo AllDone Else AllArgs.Add arg24
    If arg25 Is Nothing Then GoTo AllDone Else AllArgs.Add arg25
    If arg26 Is Nothing Then GoTo AllDone Else AllArgs.Add arg26
    If arg27 Is Nothing Then GoTo AllDone Else AllArgs.Add arg27
    If arg28 Is Nothing Then GoTo AllDone Else AllArgs.Add arg28
    If arg29 Is Nothing Then GoTo AllDone Else AllArgs.Add arg29
AllDone:

    Dim I As Long
    Dim J As Long
    
    Dim N As Long
    Dim M As Long
    
    N = AllArgs.Count
    M = -1

    Set zip = New Collection
    For I = 1 To N
        If M = -1 Then
            M = AllArgs.Item(I).Count
        ElseIf AllArgs.Item(I).Count < M Then
            M = AllArgs.Item(I).Count
        End If
    Next I

    Dim ICol As Collection
    For I = 1 To M
        Set ICol = New Collection
        For J = 1 To N
            ICol.Add AllArgs.Item(J).Item(I)
        Next J
        zip.Add ICol
    Next I
End Function




'************"MiscString"


Function randomString(length)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

