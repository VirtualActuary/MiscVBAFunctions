Attribute VB_Name = "MiscGetUniqueItems"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Private Sub TestGetUniqueItems()
    Dim arr(3) As Variant
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "b"
    Debug.Print UBound(GetUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(arr), 1), 3 ' zero index + case sensitive
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(arr, False), 1), 2 ' zero index + case insensitive
    
    arr(0) = 1: arr(1) = 2: arr(2) = 3: arr(3) = 2
    Debug.Print UBound(GetUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = 1: arr(1) = 1: arr(2) = "a": arr(3) = "a"
    Debug.Print UBound(GetUniqueItems(arr), 1), 1 ' zero index
    
End Sub

Public Function GetUniqueItems(arr() As Variant, _
            Optional CaseSensitive As Boolean = True) As Variant
    'Return an array with unique values from the input array.
    '
    ' Args:
    '   arr: Array with potential duplicate entries.
    '   CaseSensitive: If true, the duplicate checks will be case sensitive.
    '
    ' Returns:
    '   An array with unique entries.
    
    If ArrayLen(arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim d As New Dictionary
        If Not CaseSensitive Then
            d.CompareMode = TextCompare
        End If
        
        Dim I As Long
        For I = LBound(arr) To UBound(arr)
            If Not d.Exists(arr(I)) Then
                d.Add arr(I), arr(I)
            End If
        Next
        
        GetUniqueItems = d.Keys()
    End If
End Function


' Returns the number of elements in an array for a given dimension.
Private Function ArrayLen(arr As Variant, _
    Optional dimNum As Integer = 1) As Long
    
    If IsEmpty(arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function


