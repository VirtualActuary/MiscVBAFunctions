Attribute VB_Name = "MiscDataStructures"
Option Explicit


Function EnsureUniqueKey(C As Variant, Key As String, Optional Depth As Integer)
    ' Check if the key input exists in the input collection.
    ' Generate key unique to the keys in the collection.
    '
    ' Args:
    '   C: Input Collection
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
    
    If hasKey(C, EnsureUniqueKey) Then
        EnsureUniqueKey = EnsureUniqueKey(C, Key, Depth + 1)
    End If
End Function
