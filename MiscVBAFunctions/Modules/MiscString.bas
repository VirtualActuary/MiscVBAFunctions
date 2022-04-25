Attribute VB_Name = "MiscString"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function randomString(length As Variant)
    ' Create a random string containing hex characters only.
    ' (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F)
    '
    ' Args:
    '   length: Number of characters that the string must have.
    '
    ' Returns:
    '   The Random string.
    
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

