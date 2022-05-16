Attribute VB_Name = "MiscDictionary"
' This module contains functions to retrieve values given a
' key in a dictionary


Option Explicit
'@IgnoreModule ImplicitByRefModifier

Private Sub testDictget()

    Dim d As Dictionary
    Set d = dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(d, "a"), 2 ' returns 2
    Debug.Print dictget(d, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(d, "c", vbNullString), vbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(d, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Sub


Public Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
    ' Return the entry in the input Dictionary at the given key. If the given key doesn't exist,
    ' the default value is returned if it's not empty. Else an error is raised.
    '
    ' Args:
    '   d: Dictionary to read the value from...
    '   key: The key value that gets used to return the input Dictionary's value with the matching key.
    '   default: The value that must be returned if the key doesn't exist in the Dictionary.
    '
    ' Returns:
    '   The Dictionary's entry or the default value.
    
    If d.Exists(key) Then
        assign dictget, d.Item(key)
        
    ElseIf Not IsEmpty(default) Then
        assign dictget, default
        
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & key & "` "
            errmsg = errmsg & "not in dictionary"
        On Error GoTo 0
        
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, errmsg)
    End If
End Function
