Attribute VB_Name = "MiscEnsureDictIUtil"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function EnsureDictI(Container As Variant) As Object
    ' Convert all Dicts in an object to case insensitive Dicts.
    ' The object can only contain Dicts and Collections.
    '
    ' Args:
    '   Container: Object that potentially contains Dicts.
    '
    ' Returns:
    '   A dict or Collection that potentially contains Dicts and/or Collections.
    
    Dim Key As Variant
    Dim Item As Variant
    
    If TypeOf Container Is Collection Then
        Dim C As Collection
        Set C = New Collection
        
        For Each Item In Container
            If TypeOf Item Is Collection Or TypeOf Item Is Dictionary Then
                C.Add EnsureDictI(Item)
            Else
                C.Add Item
            End If
        Next Item
        
        Set EnsureDictI = C
        
    ElseIf TypeOf Container Is Dictionary Then
        Dim D As Dictionary
        Set D = New Dictionary
        D.CompareMode = TextCompare
        
        For Each Key In Container.Keys
            If TypeOf Container.Item(Key) Is Collection Or TypeOf Container.Item(Key) Is Dictionary Then
                D.Add Key, EnsureDictI(Container.Item(Key))
            Else
                D.Add Key, Container.Item(Key)
            End If
        Next Key
        
        Set EnsureDictI = D
    Else
    
        Dim Errmsg As String
        Errmsg = "ConvertToTextCompare only supports type 'Dictionary' and 'Collection'"
        On Error Resume Next: Errmsg = Errmsg & ". Got type '" & TypeName(Container) & "'": On Error GoTo 0
        Err.Raise 5, , Errmsg
        
    End If
End Function
