Attribute VB_Name = "MiscString"
Option Explicit

Function randomString(length)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

