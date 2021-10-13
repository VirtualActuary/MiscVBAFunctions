Attribute VB_Name = "MiscDictionary"
Option Explicit

Private Function testDictget()

    Dim d As Dictionary
    Set d = dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(d, "a"), 2 ' returns 2
    Debug.Print dictget(d, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(d, "c", ""), "" ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(d, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Function


Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
        
    If d.Exists(key) Then
    
        If IsObject(d.Item(key)) Then
            Set dictget = d.Item(key) 'Object
        Else
            dictget = d.Item(key) 'Variant
        End If
        
    ElseIf Not IsEmpty(default) Then
        If IsObject(default) Then
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
