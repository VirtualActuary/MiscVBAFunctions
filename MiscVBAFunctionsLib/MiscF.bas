Attribute VB_Name = "MiscF"
Option Explicit

'************"MiscCollectionCreate"


Function col(ParamArray Args() As Variant) As Collection
    Set col = New Collection
    Dim I As Long

    For I = LBound(Args) To UBound(Args)
        col.Add Args(I)
    Next

End Function


Function zip(ParamArray Args() As Variant) As Collection
    Dim I As Long
    Dim J As Long
    
    Dim N As Long
    Dim M As Long
    

    M = -1
    For I = LBound(Args) To UBound(Args)
        If M = -1 Then
            M = Args(I).Count
        ElseIf Args(I).Count < M Then
            M = Args(I).Count
        End If
    Next I

    Set zip = New Collection
    Dim ICol As Collection
    For I = 1 To M
        Set ICol = New Collection
        For J = LBound(Args) To UBound(Args)
            ICol.Add Args(J).Item(I)
        Next J
        zip.Add ICol
    Next I
End Function




'************"MiscDictionary"



Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
    Dim defType As Integer
    Dim itemType As Integer
    
    defType = 2 '2=Object
    On Error Resume Next: defType = -(default <> Empty): On Error GoTo 0 ' 0=Empty, 1=Variant
    
    If d.Exists(key) Then
    
        itemType = 2 '2=Object
        On Error Resume Next: itemType = -(d.Item(key) <> Empty): On Error GoTo 0  ' 0=Empty, 1=Variant
    
        If itemType = 2 Then
            Set dictget = d.Item(key) 'Object
        Else
            dictget = d.Item(key) 'Variant
        End If
        
    ElseIf defType <> 0 Then
        If defType = 2 Then
            Set dictget = default 'Object
        Else
            dictget = default  'Variant
        End If
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & key & "` "
            errmsg = errmsg & "not in dictionary"
        On Error GoTo 0
        
        Err.Raise 9, , errmsg
    End If
End Function

'************"MiscDictionaryCreate"


Function dict(ParamArray Args() As Variant) As Dictionary
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


Function dicti(ParamArray Args() As Variant) As Dictionary
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


'************"MiscHasKey"


Public Function hasKey(Container, key As Variant) As Boolean
    hasKey = True
    If TypeOf Container Is Collection Then
        On Error GoTo noKey
        TypeName Container(key)
        Exit Function
noKey:
        hasKey = False
    Else
        'We expect keyable VBA objects to have .Exists methods
        hasKey = Container.Exists(key)
        Exit Function
    End If
End Function

'************"MiscString"


Function randomString(length)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function


'************"MiscTables"


Function HasLO(Name As String, Optional WB As Workbook) As Boolean

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet, LO As ListObject
    
    For Each WS In WB.Sheets
        For Each LO In WS.ListObjects
            If Name = LO.Name Then
                HasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    HasLO = False

End Function


' get list object only using it's name from within a workbook
Function GetLO(Name As String, Optional WB As Workbook) As ListObject

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet, LO As ListObject
    
    For Each WS In WB.Sheets
        For Each LO In WS.ListObjects
            If Name = LO.Name Then
                Set GetLO = LO
                Exit Function
            End If
        Next LO
    Next WS
    
    If GetLO Is Nothing Then
        ' 9: Subscript out of range
        Err.Raise 9, , "List object '" & Name & "' not found in workbook '" & WB.Name & "'"
    End If

End Function
