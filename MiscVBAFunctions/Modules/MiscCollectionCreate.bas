Attribute VB_Name = "MiscCollectionCreate"
Option Explicit

Function col(ParamArray Args() As Variant) As Collection
    Set col = New Collection
    Dim i As Long

    For i = LBound(Args) To UBound(Args)
        col.Add Args(i)
    Next

End Function


Function zip(ParamArray Args() As Variant) As Collection
    Dim i As Long
    Dim J As Long
    
    Dim N As Long
    Dim M As Long
    

    M = -1
    For i = LBound(Args) To UBound(Args)
        If M = -1 Then
            M = Args(i).Count
        ElseIf Args(i).Count < M Then
            M = Args(i).Count
        End If
    Next i

    Set zip = New Collection
    Dim ICol As Collection
    For i = 1 To M
        Set ICol = New Collection
        For J = LBound(Args) To UBound(Args)
            ICol.Add Args(J).Item(i)
        Next J
        zip.Add ICol
    Next i
End Function



