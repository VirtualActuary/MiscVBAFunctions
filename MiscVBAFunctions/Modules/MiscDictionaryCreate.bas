Attribute VB_Name = "MiscDictionaryCreate"
Option Explicit

Public Function dict(ParamArray Args() As Variant) As Dictionary
    'Case sensitive dictionary
    
    Dim errmsg As String
    Set dict = New Dictionary
    
    Dim I As Long
    Dim Cnt As Long
    Cnt = 0
    For I = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If I + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(I) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dict.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function


Public Function dicti(ParamArray Args() As Variant) As Dictionary
    'Case insensitive dictionary
    
    Dim errmsg As String
    Set dicti = New Dictionary
    dicti.CompareMode = TextCompare
    
    Dim I As Long
    Dim Cnt As Long
    Cnt = 0
    For I = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If I + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(I) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dicti.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function

