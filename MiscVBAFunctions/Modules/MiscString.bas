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


Public Function IsString(ByVal v As Variant) As Boolean
    ' Test is a Variant is of type String
    '
    ' Args:
    '   v: Variant variable to test
    '
    ' Returns
    '   Boolean indication if the input is a String
    IsString = False
    
    Dim s As String
    
    ' If not an object (s=v error test) then test if not numeric
    On Error Resume Next
        s = v
        If Err.Number = 0 Then
            If Not IsNumeric(v) Then IsString = True
        End If
End Function
