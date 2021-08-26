Attribute VB_Name = "MiscDictionaryCreate"
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
    
    Dim notEmpty As Boolean
    Dim errmsg As String
    errmsg = "Dict construction is missing a pair"
    Set dict = New Dictionary
    
    '#Pattern for arg1,arg2 extended via
    'for i in range(1,30,2):
    '    print("\n", str_in.replace("arg1", "___").replace("arg2", f"arg{i+1}").replace("___", f"arg{i}"))


    ' Checking ArgX for emptyness fails on some objects like Dictionary, the error-resumption workaround handles those cases
    notEmpty = True: On Error Resume Next: notEmpty = (arg1 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg2 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg1 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg1, arg2
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg3 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg4 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg3 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg3, arg4
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg5 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg6 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg5 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg5, arg6
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg7 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg8 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg7 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg7, arg8
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg9 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg10 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg9 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg9, arg10
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg11 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg12 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg11 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg11, arg12
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg13 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg14 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg13 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg13, arg14
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg15 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg16 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg15 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg15, arg16
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg17 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg18 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg17 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg17, arg18
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg19 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg20 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg19 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg19, arg20
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg21 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg22 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg21 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg21, arg22
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg23 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg24 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg23 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg23, arg24
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg25 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg26 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg25 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg25, arg26
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg27 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg28 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg27 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg27, arg28
    Else
        Exit Function
    End If
    
    notEmpty = True: On Error Resume Next: notEmpty = (arg29 <> Empty): On Error GoTo 0
    If notEmpty Then
       notEmpty = True: On Error Resume Next: notEmpty = (arg30 <> Empty): On Error GoTo 0
        If Not notEmpty Then
            On Error Resume Next: errmsg = errmsg & " for key `" & arg29 & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
    
        dict.Add arg29, arg30
    Else
        Exit Function
    End If

End Function

