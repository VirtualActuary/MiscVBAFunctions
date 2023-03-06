Attribute VB_Name = "MiscDataStructures"
Option Explicit


Function EnsureUniqueKey(C As Variant, Key As String, Optional Depth As Long)
    ' Check if the key input exists in the input collection\Dict.
    ' Generate key unique to the keys in the collection\Dict.
    '
    ' Args:
    '   C: Input Collection\Dict
    '   Key: Input to generate a unique key from
    '   Depth: Current integer depth
    '
    ' Returns:
    '   A unique key
    
    If Depth = 0 Then
        EnsureUniqueKey = Key
    Else
        EnsureUniqueKey = Key & Depth
    End If
    
    If HasKey(C, EnsureUniqueKey) Then
        EnsureUniqueKey = EnsureUniqueKey(C, Key, Depth + 1)
    End If
End Function

