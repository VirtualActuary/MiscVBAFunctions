Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function min(ByVal col As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    min = col(1)
    
    For Each Entry In col
        If Entry < min Then
            min = Entry
        End If
    Next Entry
    
    
    
End Function

Public Function max(ByVal col As Collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    max = col(1)
    Dim Entry As Variant
    
    For Each Entry In col
        If Entry > max Then
            max = Entry
        End If
    Next Entry

End Function

Public Function mean(ByVal col As Collection) As Variant
    ' Returns the mean value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The mean value of the collection.
    
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If

    mean = 0
    Dim Entry As Variant
    
    For Each Entry In col
        mean = mean + Entry
    Next Entry
    
    mean = mean / col.Count
    
End Function


Public Function IsValueInCollection(col As Collection, val As Variant, Optional CaseSensitive As Boolean = False) As Boolean
    ' Check if a value exists in the input Collection.
    '
    ' Args:
    '   col: Collection that potentially contains val
    '   val: The value to check for.
    '   CaseSensitive: Boolean entry to indicate if the comparison must be case sensitive.
    '
    ' Returns:
    '   True if val exists in the input Collection.
    
    Dim ValI As Variant
    For Each ValI In col
        ' only check if not an object:
        If Not IsObject(ValI) Then
            If CaseSensitive Then
                IsValueInCollection = ValI = val
            Else
                IsValueInCollection = VBA.LCase(ValI) = VBA.LCase(val)
            End If
            ' exit if found
            If IsValueInCollection Then Exit Function
        End If
    Next ValI
End Function


Public Sub ConcatCollections(ParamArray CollectionArr() As Variant)
    ' Concatenate multiple Collections, thereby manipulating the first Collection.
    ' Args:
    '   CollectionArr: Array of the input collections.
    
    Dim col As Variant
    For Each col In CollectionArr
        If Not TypeOf col Is Collection Then
            Dim errmsg As String
            errmsg = "All inputs need to be Collections"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(col) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        End If
    Next col
    
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

    Dim col As Variant
    For Each col In CollectionArr
        If Not TypeOf col Is Collection Then
            Dim errmsg As String
            errmsg = "All inputs need to be Collections"
            On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(col) & "'": On Error GoTo 0
            Err.Raise 5, , errmsg
        End If
    Next col
    
    Dim I As Long
    Dim ColResult As New Collection

    For Each col In CollectionArr
        For I = 1 To col.Count
            ColResult.Add col.Item(I)
        Next
    Next col

    Set JoinCollections = ColResult
End Function
