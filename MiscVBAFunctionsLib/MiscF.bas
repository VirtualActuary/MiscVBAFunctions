Attribute VB_Name = "MiscF"
Option Explicit

'************"Module1"


Function randomString(length)
    Dim s As String
    Dim I As Long
    
    ' Add extra length redundancy for in case hex returns short of length 6
    For I = 1 To CLng((length + 6 * 2) / 5)
        s = s & Hex(Rnd * 16777216)
    Next I
    
    s = Mid(s, 1, length)
    randomString = s
End Function
