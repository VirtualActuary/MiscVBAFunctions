Attribute VB_Name = "MiscHasKey"
' Functions to check whether a key exists in a container

Option Explicit
'@IgnoreModule ImplicitByRefModifier

Private Sub TestHasKey()

    Dim C As New Collection
    C.Add "a", "a"
    C.Add Col("x", "y", "z"), "b"
    
    Debug.Print vbLf & "*********** TestHasKey tests ***********"
    Debug.Print True, hasKey(C, "a") ' True for scalar
    Debug.Print True, hasKey(C, "b") ' True for object
    Debug.Print True, hasKey(C, "A") ' True (even though case insensitive???)

    Debug.Print True, hasKey(Workbooks, ThisWorkbook.Name) ' True for non-collection type collections
    
    Dim D As New Dictionary
    D.Add "a", "a"
    D.Add "b", Col("x", "y", "z")
    
    Debug.Print True, hasKey(D, "a") ' True for scalar
    Debug.Print True, hasKey(D, "b") ' True for object
    Debug.Print False, hasKey(D, "A") ' False - case sensitive by default
    
    Dim dObj As Object
    Set dObj = CreateObject("Scripting.Dictionary")
    
    dObj.Add "a", "a"
    dObj.Add "b", Col("x", "y", "z")
    
    Debug.Print True, hasKey(dObj, "a") ' True for scalar
    Debug.Print True, hasKey(dObj, "b") ' True for object
    Debug.Print False, hasKey(dObj, "A") ' False - case sensitive by default
    
    ' Errors
    On Error Resume Next
        Err.Number = 0
        hasKey ThisWorkbook, "A"
        Debug.Print 9, Err.Number ' WorkBook doesn't have keys/items
        
        Err.Number = 0
        hasKey 5, "A"
        Debug.Print 9, Err.Number ' Variant doesn't have keys/items
    On Error GoTo 0

End Sub


Public Function hasKey(Container As Variant, Key As Variant) As Boolean
    ' Checks whether a key exists in an existing container
    ' The container can be a `Collection`, `Dictionary` or any
    ' built-in Dictionary-like object. For example `ThisWorkbook.Sheets`
    '
    ' Args:
    '     Container: The container in which to look for the key
    '     key: The key to look for in the container
    '
    ' Returns:
    '     True for success, False otherwise
    
    
    Dim ErrX As Integer
    Dim hasKeyFlag As Boolean
    Dim emptyFlag As Boolean
    
    ' First try .HasKey method on the object
    On Error Resume Next
        Err.Number = 0
        hasKeyFlag = Container.Exists(Key)
        ErrX = Err.Number
    On Error GoTo 0
    If ErrX = 0 Then
        hasKey = hasKeyFlag
        Exit Function
    End If
    
    
    ' Then test with .Item method
    emptyFlag = False
    On Error Resume Next
        Err.Number = 0
        emptyFlag = TypeName(Container.Item(Key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via .Item
        hasKey = Not emptyFlag
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 Then ' Retrieval Error, but .Item is correct access method stil. 424: Method not exist; 438: Compilation error
        hasKey = False
        Exit Function
    End If
    
    
    ' Then test with bracketed access, like ()
    emptyFlag = False
    On Error Resume Next
        Err.Number = 0
        emptyFlag = TypeName(Container(Key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via ()
        hasKey = Not emptyFlag
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 And ErrX <> 13 Then ' Retrieval Error, but () is correct access method stil. 424: Method not exist; 438: Compilation error; 13: Variant bracketed ()
        hasKey = False
        Exit Function
    End If

    
    Dim errmsg As String
    On Error Resume Next
        errmsg = "Object"
        errmsg = errmsg & " of type '" & TypeName(Container) & "'"
        errmsg = errmsg & " have neither '.Exists' method, nor bracketed indexing '()', nor '.Item' method"
    On Error GoTo 0
    Err.Raise 9, , errmsg
    
End Function

