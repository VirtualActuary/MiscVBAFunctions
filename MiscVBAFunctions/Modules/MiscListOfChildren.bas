Attribute VB_Name = "MiscListOfChildren"
Option Explicit


Function GetListOfChildren(Depths As Collection, Optional SearchDown As Boolean = True) As Collection
    ' Given a collection of depths (integers), returns the position of the immediate
    ' children for each position
    '
    ' Args:
    '    Depths: a collection of the depths for each element from which the children can be inferred
    '    SearchDown: whether to look downwards or upwards for children. Defaults to look downwards
    '
    ' Returns:
    '    Collection for each item that contains the position of its children
    
    Dim ListOfChildren As New Collection
    
    Dim C As Long
    ' Create empty collections for the output
    For C = 1 To Depths.Count
        ListOfChildren.Add New Collection
    Next C
    
    Dim I As Long
    Dim Idx As Long
    
    If SearchDown Then ' look down / forward in list
        For I = 1 To Depths.Count
            Idx = I
            Do While Idx < Depths.Count ' last one doesn't have children by definition
                Idx = Idx + 1 ' move down / forward in list
                If Depths(Idx) <= Depths(I) Then
                    Exit Do
                End If
                
                If Depths(Idx) = Depths(I) + 1 Then
                    ListOfChildren(I).Add Idx
                End If
            Loop
            
        Next I
    Else ' look up / backwards in list
        For I = 1 To Depths.Count
            Idx = I
            Do While Idx > 1 ' first one doesn't have children by definition
                Idx = Idx - 1 ' move up / backwards in list
                If Depths(Idx) <= Depths(I) Then
                    Exit Do
                End If
                
                If Depths(Idx) = Depths(I) + 1 Then
                    ListOfChildren(I).Add Idx
                End If
            Loop
            
        Next I
    End If
    
    Set GetListOfChildren = ListOfChildren
    
End Function



