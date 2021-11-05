Attribute VB_Name = "MiscDictionaryCreate"
Option Explicit

Function dict(ParamArray Args() As Variant) As Dictionary
    'Case sensitive dictionary
    
    Dim errmsg As String
    Set dict = New Dictionary
    
    Dim i As Long
    Dim Cnt As Long
    Cnt = 0
    For i = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If i + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(i) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dict.Add Args(i), Args(i + 1)
Cont:
    Next i

End Function


Function dicti(ParamArray Args() As Variant) As Dictionary
    'Case insensitive dictionary
    
    Dim errmsg As String
    Set dicti = New Dictionary
    dicti.CompareMode = TextCompare
    
    Dim i As Long
    Dim Cnt As Long
    Cnt = 0
    For i = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If i + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(i) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dicti.Add Args(i), Args(i + 1)
Cont:
    Next i

End Function

