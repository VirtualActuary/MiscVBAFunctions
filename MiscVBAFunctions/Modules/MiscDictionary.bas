Attribute VB_Name = "MiscDictionary"
' This module contains functions to retrieve values given a
' key in a dictionary


Option Explicit
'@IgnoreModule ImplicitByRefModifier

Private Sub TestDictget()

    Dim D As Dictionary
    Set D = Dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print Dictget(D, "a"), 2 ' returns 2
    Debug.Print Dictget(D, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print Dictget(D, "c", VbNullString), VbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print Dictget(D, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Sub


Public Function Dictget(D As Dictionary, Key As Variant, Optional Default As Variant = Empty) As Variant
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
        Assign Dictget, D.Item(Key)
        
    ElseIf Not IsEmpty(Default) Then
        Assign Dictget, Default
        
    Else
        Dim Errmsg As String
        On Error Resume Next
            Errmsg = "Key "
            Errmsg = Errmsg & "`" & Key & "` "
            Errmsg = Errmsg & "not in dictionary"
        On Error GoTo 0
        
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, Errmsg)
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
    
    Dim DictCompareMode As Long
    DictCompareMode = -999  ' This is a placeholder value
    Dim Dict As Variant
    For Each Dict In Dicts
        If Not TypeOf Dict Is Dictionary Then
            Dim Errmsg As String
            Errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: Errmsg = Errmsg & ". Got type '" & TypeName(Dict) & "'": On Error GoTo 0
            Err.Raise 5, , Errmsg
        Else
            If DictCompareMode <> -999 And DictCompareMode <> Dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            DictCompareMode = Dict.CompareMode
        
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
    
    Dim DictCompareMode As Long
    DictCompareMode = -999  ' This is a placeholder value
    Dim Dict As Variant
    For Each Dict In Dicts
        If Not TypeOf Dict Is Dictionary Then
            Dim Errmsg As String
            Errmsg = "All inputs need to be Dictionaries"
            On Error Resume Next: Errmsg = Errmsg & ". Got type '" & TypeName(Dict) & "'": On Error GoTo 0
            Err.Raise 5, , Errmsg
        Else
            If DictCompareMode <> -999 And DictCompareMode <> Dict.CompareMode Then
                Err.Raise -987, , "CompareMode of all Dictionaries aren't the same."
            End If
            DictCompareMode = Dict.CompareMode
        
        End If
    Next Dict
   
    Dim D As New Dictionary
    D.CompareMode = DictCompareMode
    Dim Key As Variant

    For Each Dict In Dicts
        For Each Key In Dict.Keys
            D(Key) = Dict.Item(Key)
        Next Key

    Next Dict

    Set JoinDicts = D

End Function


Function DictToCollection(D As Dictionary) As Collection
    ' Create a keyed collection from a dictionary
    '
    ' Collections may be used like dictionaries, but it's not convenient, because you can't get a list of available keys.
    ' You probably do not want to do this. Rather use a dictionary!
    '
    ' This might be useful in some backwards-compatibility situations.
    '
    ' Args:
    '   Args: List of keys and values that gets inserted into the Collection.
    '         All uneven entries are the keys and all even entries are the values for the matching keys.
    '
    ' Returns:
    '   The Collection
    
    Dim Key As Variant
    Set DictToCollection = New Collection
    For Each Key In D.Keys
        DictToCollection.Add D.Item(Key), Key
    Next Key
End Function
