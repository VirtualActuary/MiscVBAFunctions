Attribute VB_Name = "MiscCollectionCreate"
Option Explicit

Public Function Col(ParamArray Args() As Variant) As Collection
    ' Create a Collection from a list of entries.
    '
    ' Args:
    '   Args: list of entries that gets inserted into the Collection
    '
    ' Returns:
    '   Collection with the arguement values inserted.
    
    Set Col = New Collection
    Dim I As Long

    For I = LBound(Args) To UBound(Args)
        Col.Add Args(I)
    Next

End Function


Public Function Zip(ParamArray Args() As Variant) As Collection
    ' Standard zip function. Takes multiple Collections as an argument and
    ' group the matching index entries of each Collection into a new Collection.
    '
    ' Args:
    '   Args: Multiple Collections that gets grouped by index number.
    '
    ' Returns:
    '   A collection of collections containing the grouped entries.
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    
    M = -1
    For I = LBound(Args) To UBound(Args)
        If M = -1 Then
            M = Args(I).Count
        ElseIf Args(I).Count < M Then
            M = Args(I).Count
        End If
    Next I

    Set Zip = New Collection
    Dim ICol As Collection
    For I = 1 To M
        Set ICol = New Collection
        For J = LBound(Args) To UBound(Args)
            ICol.Add Args(J).Item(I)
        Next J
        Zip.Add ICol
    Next I
End Function



