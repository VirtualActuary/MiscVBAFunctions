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
    
    Dim key As Variant
    Dim Item As Variant
    
    If TypeOf Container Is Collection Then
        Dim c As Collection
        Set c = New Collection
        
        For Each Item In Container
            If TypeOf Item Is Collection Or TypeOf Item Is Dictionary Then
                c.Add EnsureDictI(Item)
            Else
                c.Add Item
            End If
        Next Item
        
        Set EnsureDictI = c
        
    ElseIf TypeOf Container Is Dictionary Then
        Dim d As Dictionary
        Set d = New Dictionary
        d.CompareMode = TextCompare
        
        For Each key In Container.Keys
            If TypeOf Container.Item(key) Is Collection Or TypeOf Container.Item(key) Is Dictionary Then
                d.Add key, EnsureDictI(Container.Item(key))
            Else
                d.Add key, Container.Item(key)
            End If
        Next key
        
        Set EnsureDictI = d
    Else
    
        Dim errmsg As String
        errmsg = "ConvertToTextCompare only supports type 'Dictionary' and 'Collection'"
        On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Container) & "'": On Error GoTo 0
        Err.Raise 5, , errmsg
        
    End If
End Function
