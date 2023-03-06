Attribute VB_Name = "MiscHasKey"
' Functions to check whether a key exists in a container

Option Explicit
'@IgnoreModule ImplicitByRefModifier

Private Sub TestHasKey()

    Dim C As New Collection
    C.Add "a", "a"
    C.Add Col("x", "y", "z"), "b"
    
    Debug.Print VbLf & "*********** TestHasKey tests ***********"
    Debug.Print True, HasKey(C, "a") ' True for scalar
    Debug.Print True, HasKey(C, "b") ' True for object
    Debug.Print True, HasKey(C, "A") ' True (even though case insensitive???)

    Debug.Print True, HasKey(Workbooks, ThisWorkbook.Name) ' True for non-collection type collections
    
    Dim D As New Dictionary
    D.Add "a", "a"
    D.Add "b", Col("x", "y", "z")
    
    Debug.Print True, HasKey(D, "a") ' True for scalar
    Debug.Print True, HasKey(D, "b") ' True for object
    Debug.Print False, HasKey(D, "A") ' False - case sensitive by default
    
    Dim DObj As Object
    Set DObj = CreateObject("Scripting.Dictionary")
    
    DObj.Add "a", "a"
    DObj.Add "b", Col("x", "y", "z")
    
    Debug.Print True, HasKey(DObj, "a") ' True for scalar
    Debug.Print True, HasKey(DObj, "b") ' True for object
    Debug.Print False, HasKey(DObj, "A") ' False - case sensitive by default
    
    ' Errors
    On Error Resume Next
        Err.Number = 0
        HasKey ThisWorkbook, "A"
        Debug.Print 9, Err.Number ' WorkBook doesn't have keys/items
        
        Err.Number = 0
        HasKey 5, "A"
        Debug.Print 9, Err.Number ' Variant doesn't have keys/items
    On Error GoTo 0

End Sub


Public Function HasKey(Container As Variant, Key As Variant) As Boolean
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
    Dim HasKeyFlag As Boolean
    Dim EmptyFlag As Boolean
    
    ' First try .HasKey method on the object
    On Error Resume Next
        Err.Number = 0
        HasKeyFlag = Container.Exists(Key)
        ErrX = Err.Number
    On Error GoTo 0
    If ErrX = 0 Then
        HasKey = HasKeyFlag
        Exit Function
    End If
    
    
    ' Then test with .Item method
    EmptyFlag = False
    On Error Resume Next
        Err.Number = 0
        EmptyFlag = TypeName(Container.Item(Key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via .Item
        HasKey = Not EmptyFlag
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 Then ' Retrieval Error, but .Item is correct access method stil. 424: Method not exist; 438: Compilation error
        HasKey = False
        Exit Function
    End If
    
    
    ' Then test with bracketed access, like ()
    EmptyFlag = False
    On Error Resume Next
        Err.Number = 0
        EmptyFlag = TypeName(Container(Key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via ()
        HasKey = Not EmptyFlag
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 And ErrX <> 13 Then ' Retrieval Error, but () is correct access method stil. 424: Method not exist; 438: Compilation error; 13: Variant bracketed ()
        HasKey = False
        Exit Function
    End If

    
    Dim Errmsg As String
    On Error Resume Next
        Errmsg = "Object"
        Errmsg = Errmsg & " of type '" & TypeName(Container) & "'"
        Errmsg = Errmsg & " have neither '.Exists' method, nor bracketed indexing '()', nor '.Item' method"
    On Error GoTo 0
    Err.Raise 9, , Errmsg
    
End Function

