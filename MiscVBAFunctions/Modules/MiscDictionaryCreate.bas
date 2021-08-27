Attribute VB_Name = "MiscDictionaryCreate"
Option Explicit

Function dict(ParamArray Args() As Variant) As Dictionary
    
    Dim errmsg As String
    Set dict = New Dictionary
    
    Dim I As Long
    Dim II As Long
    II = LBound(Args)
    For I = LBound(Args) To Round(UBound(Args) + 0.999 / 2) - 1
        If II + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(II) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dict.Add Args(II), Args(II + 1)
        II = II + 2
    Next I

End Function

