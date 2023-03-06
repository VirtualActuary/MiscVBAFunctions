Attribute VB_Name = "MiscCollectionSort"
Option Explicit


Private Sub TestBubbleSort()
    Dim C As Collection
    Set C = Col("variables10", "variables", "variables2", "variables_10", "variables_2")
    Set C = BubbleSort(Coll)
    
    Debug.Print C(1), "variables"
    Debug.Print C(2), "variables10" ' :/
    Debug.Print C(3), "variables2" ' :/
    Debug.Print C(4), "variables_10" ' :/
    Debug.Print C(5), "variables_2" ' :/
    
End Sub


Public Function BubbleSort(Coll As Collection) As Collection
    
    ' from: https://github.com/austinleedavis/VBA-utilities/blob/f23f1096d8df0dfdc740e5a3bec36525d61a3ffc/Collections.bas#L73
    ' this is an easy implementation but a slow sorting algorithm.
    ' do not use for large collections.
    '
    ' Args:
    '   coll: Unsorted Collection.
    '
    ' Returns:
    '   Sorted Collection
    
    Dim SortedColl As Collection
    Set SortedColl = New Collection
    Dim VItm As Variant
    ' copy the collection"
    For Each VItm In Coll
        SortedColl.Add VItm
    Next VItm

    Dim I As Long, J As Long
    Dim VTemp As Variant

    'Two loops to bubble sort
    For I = 1 To SortedColl.Count - 1
        For J = I + 1 To SortedColl.Count
            If SortedColl(I) > SortedColl(J) Then ' 1 = I is larger than J
                'store the lesser item
               Assign VTemp, SortedColl(J) ' assign
                'remove the lesser item
               SortedColl.Remove J
                're-add the lesser item before the greater Item
               SortedColl.Add VTemp, , I
            End If
        Next J
    Next I
    
    Set BubbleSort = SortedColl
    
End Function

