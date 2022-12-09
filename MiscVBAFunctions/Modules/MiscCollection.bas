Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function min(ByVal Col As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    min = Col(1)
    
    For Each Entry In Col
        If Entry < min Then
            min = Entry
        End If
    Next Entry
    
    
    
End Function

Public Function max(ByVal Col As Collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    max = Col(1)
    Dim Entry As Variant
    
    For Each Entry In Col
        If Entry > max Then
            max = Entry
        End If
    Next Entry

End Function

Public Function mean(ByVal Col As Collection) As Variant
    ' Returns the mean value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The mean value of the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If

    mean = 0
    Dim Entry As Variant
    
    For Each Entry In Col
        mean = mean + Entry
    Next Entry
    
    mean = mean / Col.Count
    
End Function


Public Function IsValueInCollection(ColInput As Collection, Val As Variant, Optional CaseSensitive As Boolean = False) As Boolean
    ' Check if a value exists in the input Collection.
    '
    ' Args:
    '   ColInput: Collection that potentially contains val
    '   val: The value to check for.
    '   CaseSensitive: Boolean entry to indicate if the comparison must be case sensitive.
    '
    ' Returns:
    '   True if val exists in the input Collection.
    
    Dim ValI As Variant
    For Each ValI In ColInput
        ' only check if not an object:
        If Not IsObject(ValI) Then
            If CaseSensitive Then
                IsValueInCollection = ValI = Val
            Else
                IsValueInCollection = VBA.LCase(ValI) = VBA.LCase(Val)
            End If
            ' exit if found
            If IsValueInCollection Then Exit Function
        End If
    Next ValI
End Function


Public Function indexOf(ByVal Col1 As Collection, ByVal Item As Variant) As Long
    'https://stackoverflow.com/questions/28985579/retrieve-the-index-of-an-object-stored-in-a-collection-using-its-key-vba
    'returns index of item if found, returns 0 if not found
    '
    ' Args:
    '   Col1: Collection that potentially contains the item.
    '   item: Item to find the index of.
    '
    ' Returns:
    '   The index of the item in the Collection.
    
    Dim I As Long
    For I = 1 To Col1.Count

        If Col1(I) = Item Then
            indexOf = I
            Exit Function
        End If
        
    Next
End Function


Public Sub ConcatCollections(ParamArray CollectionArr() As Variant)
    ' Concatenate multiple Collections, thereby manipulating the first Collection.
    ' Args:
    '   CollectionArr: Array of the input collections.
    
    Dim Col As Variant
    For Each Col In CollectionArr
        If Not TypeOf Col Is Collection Then
            Dim errmsg As String
            errmsg = "All inputs need to be Collections"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Col) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        End If
    Next Col
    
    Dim I As Long
    Dim J As Long
    For J = 1 To UBound(CollectionArr)
        For I = 1 To CollectionArr(J).Count
            CollectionArr(0).Add CollectionArr(J).Item(I)
        Next
    Next

End Sub


Public Function JoinCollections(ParamArray CollectionArr()) As Collection
    ' Joins multiple Collections and returns the result.
    ' None of the inputs get manipulated.
    '
    ' Args:
    '   CollectionArr: Array of the input collections.
    '
    ' Returns:
    '   Returns a new Collection of the joined Collections.

    Dim Col As Variant
    For Each Col In CollectionArr
        If Not TypeOf Col Is Collection Then
            Dim errmsg As String
            errmsg = "All inputs need to be Collections"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Col) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        End If
    Next Col
    
    Dim I As Long
    Dim ColResult As New Collection

    For Each Col In CollectionArr
        For I = 1 To Col.Count
            ColResult.Add Col.Item(I)
        Next
    Next Col

    Set JoinCollections = ColResult
End Function


Function CollectionToArray(C As Collection) As Variant
    ' Create an Array from the content of a Collection.
    ' Ignore the Collection's keys.
    '
    ' Args:
    '   C: Input Collection
    '
    ' Returns:
    '   An array with the same content as the input collection.
    
    If C.Count < 1 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    Dim Result As Variant
    Dim I As Long
    ReDim Result(C.Count - 1)
    
    For I = 1 To C.Count
        Result(I - 1) = C(I)
    Next I
    
    CollectionToArray = Result
End Function
