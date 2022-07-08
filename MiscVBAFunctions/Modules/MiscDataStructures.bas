Attribute VB_Name = "MiscDataStructures"
Option Explicit


Public Sub ConcatContainers(DataStructure1 As Variant, DataStructure2 As Variant)
    If TypeOf DataStructure1 Is Collection And TypeOf DataStructure2 Is Collection Then
        Dim I As Long
        For I = 1 To DataStructure2.Count
            DataStructure1.Add DataStructure2.Item(I)
        Next
    ElseIf TypeOf DataStructure1 Is Dictionary And TypeOf DataStructure2 Is Dictionary Then
        Dim key As Variant
        For Each key In DataStructure2.Keys
            DataStructure1.Add key, DataStructure2.Item(key)
        Next key
    Else
    
        Dim errmsg As String
        errmsg = "ConvertToTextCompare only supports type 'Dictionary and Dictionary' or 'Collection and Collection'"
        On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(DataStructure1) & "' and '" & TypeName(DataStructure1) & "'": On Error GoTo 0
        Err.Raise 5, , errmsg
        
    End If
End Sub



Public Function JoinContainers(DataStructure1 As Variant, ByVal DataStructure2 As Variant) As Variant
    If TypeOf DataStructure1 Is Collection And TypeOf DataStructure2 Is Collection Then
        Dim I As Long
        Dim DS As New Collection
        For I = 1 To DataStructure1.Count
            DS.Add DataStructure1.Item(I)
        Next
    
        For I = 1 To DataStructure2.Count
            DS.Add DataStructure2.Item(I)
        Next
        Set JoinContainers = DS
    
    ElseIf TypeOf DataStructure1 Is Dictionary And TypeOf DataStructure2 Is Dictionary Then
        Dim d As New Dictionary
        Dim key As Variant
        For Each key In DataStructure1.Keys
            d.Add key, DataStructure1.Item(key)
        Next key
        
        For Each key In DataStructure2.Keys
            d.Add key, DataStructure2.Item(key)
        Next key
        Set JoinContainers = d
    Else
    
        Dim errmsg As String
        errmsg = "ConvertToTextCompare only supports type 'Dictionary and Dictionary' or 'Collection and Collection'"
        On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(DataStructure1) & "' and '" & TypeName(DataStructure1) & "'": On Error GoTo 0
        Err.Raise 5, , errmsg
        
    End If
End Function

