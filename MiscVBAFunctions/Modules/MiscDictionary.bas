Attribute VB_Name = "MiscDictionary"
' This module contains functions to retrieve values given a
' key in a dictionary


Option Explicit
'@IgnoreModule ImplicitByRefModifier

Private Sub testDictget()

    Dim D As Dictionary
    Set D = Dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(D, "a"), 2 ' returns 2
    Debug.Print dictget(D, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(D, "c", vbNullString), vbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(D, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Sub


Public Function dictget(D As Dictionary, Key As Variant, Optional default As Variant = Empty) As Variant
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
    
    If D.Exists(Key) Then
        assign dictget, D.Item(Key)
        
    ElseIf Not IsEmpty(default) Then
        assign dictget, default
        
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & Key & "` "
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
    Dim Dict As Variant
    For Each Dict In Dicts
        If Not TypeOf Dict Is Dictionary Then
            Dim errmsg As String
            errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Dict) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        Else
            If dictCompareMode <> -999 And dictCompareMode <> Dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            dictCompareMode = Dict.CompareMode
        
        End If
    Next Dict

    Dim J As Long

    Dim Key As Variant
    For J = 1 To UBound(Dicts)
        For Each Key In Dicts(J).Keys
            Dicts(0)(Key) = Dicts(J).Item(Key)
        Next Key

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
    Dim Dict As Variant
    For Each Dict In Dicts
        If Not TypeOf Dict Is Dictionary Then
            Dim errmsg As String
            errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Dict) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        Else
            If dictCompareMode <> -999 And dictCompareMode <> Dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            dictCompareMode = Dict.CompareMode
        
        End If
    Next Dict
   
    Dim D As New Dictionary
    D.CompareMode = dictCompareMode
    Dim Key As Variant

    For Each Dict In Dicts
        For Each Key In Dict.Keys
            D(Key) = Dict.Item(Key)
        Next Key

    Next Dict

    Set JoinDicts = D

End Function


