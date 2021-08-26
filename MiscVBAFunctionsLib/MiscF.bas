Attribute VB_Name = "MiscF"
Option Explicit

'************"MiscCollectionCreate"


Function col( _
    Optional arg1 As Variant = Empty, Optional arg2 As Variant = Empty, Optional arg3 As Variant = Empty, Optional arg4 As Variant = Empty, _
    Optional arg5 As Variant = Empty, Optional arg6 As Variant = Empty, Optional arg7 As Variant = Empty, Optional arg8 As Variant = Empty, _
    Optional arg9 As Variant = Empty, Optional arg10 As Variant = Empty, Optional arg11 As Variant = Empty, Optional arg12 As Variant = Empty, _
    Optional arg13 As Variant = Empty, Optional arg14 As Variant = Empty, Optional arg15 As Variant = Empty, Optional arg16 As Variant = Empty, _
    Optional arg17 As Variant = Empty, Optional arg18 As Variant = Empty, Optional arg19 As Variant = Empty, Optional arg20 As Variant = Empty, _
    Optional arg21 As Variant = Empty, Optional arg22 As Variant = Empty, Optional arg23 As Variant = Empty, Optional arg24 As Variant = Empty, _
    Optional arg25 As Variant = Empty, Optional arg26 As Variant = Empty, Optional arg27 As Variant = Empty, Optional arg28 As Variant = Empty, _
    Optional arg29 As Variant = Empty, Optional arg30 As Variant = Empty) As Collection
    
    Dim notEmpty As Boolean
    Set col = New Collection
    
    ' Checking ArgX for emptyness fails on some objects like Dictionary, the error-resumption workaround handles those cases
    notEmpty = True: On Error Resume Next: notEmpty = (arg1 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg1 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg2 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg2 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg3 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg3 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg4 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg4 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg5 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg5 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg6 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg6 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg7 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg7 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg8 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg8 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg9 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg9 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg10 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg10 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg11 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg11 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg12 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg12 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg13 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg13 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg14 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg14 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg15 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg15 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg16 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg16 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg17 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg17 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg18 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg18 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg19 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg19 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg20 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg20 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg21 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg21 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg22 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg22 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg23 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg23 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg24 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg24 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg25 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg25 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg26 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg26 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg27 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg27 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg28 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg28 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg29 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg29 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg30 <> Empty): On Error GoTo 0: If notEmpty Then col.Add arg30 Else Exit Function
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




'************"MiscDictionary"



Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
    Dim defType As Integer
    Dim itemType As Integer
    
    defType = 2 '2=Object
    On Error Resume Next: defType = -(default <> Empty): On Error GoTo 0 ' 0=Empty, 1=Variant
    
    If d.Exists(key) Then
    
        itemType = 2 '2=Object
        On Error Resume Next: itemType = -(d.Item(key) <> Empty): On Error GoTo 0  ' 0=Empty, 1=Variant
    
        If itemType = 2 Then
            Set dictget = d.Item(key) 'Object
        Else
            dictget = d.Item(key) 'Variant
        End If
        
    ElseIf defType <> 0 Then
        If defType = 2 Then
            Set dictget = default 'Object
        Else
            dictget = default  'Variant
        End If
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & key & "` "
            errmsg = errmsg & "not in dictionary"
        On Error GoTo 0
        
        Err.Raise 9, , errmsg
    End If
End Function

'************"MiscDictionaryCreate"


Function dict( _
    Optional arg1 As Variant = Empty, Optional arg2 As Variant = Empty, Optional arg3 As Variant = Empty, Optional arg4 As Variant = Empty, _
    Optional arg5 As Variant = Empty, Optional arg6 As Variant = Empty, Optional arg7 As Variant = Empty, Optional arg8 As Variant = Empty, _
    Optional arg9 As Variant = Empty, Optional arg10 As Variant = Empty, Optional arg11 As Variant = Empty, Optional arg12 As Variant = Empty, _
    Optional arg13 As Variant = Empty, Optional arg14 As Variant = Empty, Optional arg15 As Variant = Empty, Optional arg16 As Variant = Empty, _
    Optional arg17 As Variant = Empty, Optional arg18 As Variant = Empty, Optional arg19 As Variant = Empty, Optional arg20 As Variant = Empty, _
    Optional arg21 As Variant = Empty, Optional arg22 As Variant = Empty, Optional arg23 As Variant = Empty, Optional arg24 As Variant = Empty, _
    Optional arg25 As Variant = Empty, Optional arg26 As Variant = Empty, Optional arg27 As Variant = Empty, Optional arg28 As Variant = Empty, _
    Optional arg29 As Variant = Empty, Optional arg30 As Variant = Empty) As Dictionary
    
    Dim notEmpty As Boolean
    Set dict = New Dictionary
    ' Checking ArgX for emptyness fails on some objects like Dictionary, the error-resumption workaround handles those cases
    notEmpty = True: On Error Resume Next: notEmpty = (arg1 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg1, arg2 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg3 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg3, arg4 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg5 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg5, arg6 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg7 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg7, arg8 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg9 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg9, arg10 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg11 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg11, arg12 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg13 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg13, arg14 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg15 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg15, arg16 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg17 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg17, arg18 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg19 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg19, arg20 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg21 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg21, arg22 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg23 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg23, arg24 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg25 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg25, arg26 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg27 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg27, arg28 Else Exit Function
    notEmpty = True: On Error Resume Next: notEmpty = (arg29 <> Empty): On Error GoTo 0: If notEmpty Then dict.Add arg29, arg30 Else Exit Function
End Function


'************"MiscString"


Function randomString(length)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

