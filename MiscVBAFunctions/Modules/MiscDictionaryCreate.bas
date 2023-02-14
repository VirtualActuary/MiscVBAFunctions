Attribute VB_Name = "MiscDictionaryCreate"
Option Explicit

Public Function Dict(ParamArray Args() As Variant) As Dictionary
    ' Case sensitive dictionary
    '
    ' Args:
    '   Args: List of values that gets inserted into the Dictionary.
    '         All uneven entries are the keys and all even entries are the values for the matching keys.
    '
    ' Returns:
    '   The Dictionary
    
    Dim errmsg As String
    Set Dict = New Dictionary
    
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
        
        Dict.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function


Public Function DictI(ParamArray Args() As Variant) As Dictionary
    ' Case insensitive dictionary
    '
    ' Args:
    '   Args: List of values that gets inserted into the Dictionary.
    '         All uneven entries are the keys and all even entries are the values at its matching key.
    '
    ' Returns:
    '   The case insensitive Dictionary
    
    Dim errmsg As String
    Set DictI = New Dictionary
    DictI.CompareMode = TextCompare
    
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
        
        DictI.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function

