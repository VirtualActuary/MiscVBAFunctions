Attribute VB_Name = "MiscString"
Option Explicit

Public Function randomString(length As Variant)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

