Attribute VB_Name = "MiscGetUniqueItems"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Private Sub TestGetUniqueItems()
    Dim Arr(3) As Variant
    
    Arr(0) = "a": Arr(1) = "b": Arr(2) = "c": Arr(3) = "b"
    Debug.Print UBound(GetUniqueItems(Arr), 1), 2 ' zero index
    
    Arr(0) = "a": Arr(1) = "b": Arr(2) = "c": Arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(Arr), 1), 3 ' zero index + case sensitive
    
    Arr(0) = "a": Arr(1) = "b": Arr(2) = "c": Arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(Arr, False), 1), 2 ' zero index + case insensitive
    
    Arr(0) = 1: Arr(1) = 2: Arr(2) = 3: Arr(3) = 2
    Debug.Print UBound(GetUniqueItems(Arr), 1), 2 ' zero index
    
    Arr(0) = 1: Arr(1) = 1: Arr(2) = "a": Arr(3) = "a"
    Debug.Print UBound(GetUniqueItems(Arr), 1), 1 ' zero index
    
End Sub

Public Function GetUniqueItems(Arr() As Variant, _
            Optional CaseSensitive As Boolean = True) As Variant
    'Return an array with unique values from the input array.
    '
    ' Args:
    '   arr: Array with potential duplicate entries.
    '   CaseSensitive: If true, the duplicate checks will be case sensitive.
    '
    ' Returns:
    '   An array with unique entries.
    
    If ArrayLen(Arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim D As New Dictionary
        If Not CaseSensitive Then
            D.CompareMode = TextCompare
        End If
        
        Dim I As Long
        For I = LBound(Arr) To UBound(Arr)
            If Not D.Exists(Arr(I)) Then
                D.Add Arr(I), Arr(I)
            End If
        Next
        
        GetUniqueItems = D.Keys()
    End If
End Function


' Returns the number of elements in an array for a given dimension.
Private Function ArrayLen(Arr As Variant, _
    Optional DimNum As Integer = 1) As Long
    
    If IsEmpty(Arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(Arr, DimNum) - LBound(Arr, DimNum) + 1
    End If
End Function


