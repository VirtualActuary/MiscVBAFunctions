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


Public Sub ConcatDicts(ParamArray Dicts())
    ' Concatenate Dictionaries to the first dictionary, thereby manipulating the first dictionary.
    ' Nothing gets returned since the first Dictionary gets manipulated.
    ' If not all of the inputs are Dictionaries, an error will be raised.
    ' If the CompareMode of the different dictionaries don't match, an error will be raised.
    '
    ' Args:
    '   dicts: Array of dictionaries.
    
    Dim dictCompareMode As Long
    dictCompareMode = -999  ' This is a placeholder value
    Dim dict As Variant
    For Each dict In Dicts
        If Not TypeOf dict Is Dictionary Then
            Dim errmsg As String
            errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(dict) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        Else
            If dictCompareMode <> -999 And dictCompareMode <> dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            dictCompareMode = dict.CompareMode
        
        End If
    Next dict

    Dim J As Long

    Dim key As Variant
    For J = 1 To UBound(Dicts)
        For Each key In Dicts(J).Keys
            Dicts(0)(key) = Dicts(J).Item(key)
        Next key

    Next

End Sub


Public Function JoinDicts(ParamArray Dicts()) As Dictionary
    ' Joins multiple Dictionaries and returns the result.
    ' None of the inputs get manipulated.
    ' If not all of the inputs are Dictionaries, an error will be raised.
    ' If the CompareMode of the different dictionaries don't match, an error will be raised.
    '
    ' Args:
    '   dicts: Array of the input Dictionaries.
    '
    ' Returns:
    '   Returns a new Dictionary of the joined Dictionaries.
    
    Dim dictCompareMode As Long
    dictCompareMode = -999  ' This is a placeholder value
    Dim dict As Variant
    For Each dict In Dicts
        If Not TypeOf dict Is Dictionary Then
            Dim errmsg As String
            errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(dict) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        Else
            If dictCompareMode <> -999 And dictCompareMode <> dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            dictCompareMode = dict.CompareMode
        
        End If
    Next dict
   
    Dim d As New Dictionary
    d.CompareMode = dictCompareMode
    Dim key As Variant

    For Each dict In Dicts
        For Each key In dict.Keys
            d(key) = dict.Item(key)
        Next key

    Next dict

    Set JoinDicts = d

End Function


