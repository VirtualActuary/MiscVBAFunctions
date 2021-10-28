Attribute VB_Name = "MiscCollection"
Option Explicit


Function min(ByVal col As Collection) As Variant
    
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

Function max(ByVal col As Collection) As Variant
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    max = col(1)
    
    For Each Entry In col
        If Entry > max Then
            max = Entry
        End If
    Next Entry

End Function

Function mean(ByVal col As Collection) As Variant
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If

    mean = 0
    
    For Each Entry In col
        mean = mean + Entry
    Next Entry
    
    mean = mean / col.Count
    
End Function





