Attribute VB_Name = "MiscExistsInCollection"
Option Explicit

Public Function ExistsInCollection(col, key As Variant) As Boolean
    ' if either an object or non-object exists:
    ExistsInCollection = ContainsObject(col, key) Or ContainsNonObject(col, key)
End Function


' whether an object exists in a collection
Function ContainsObject(col, key As Variant) As Boolean
    Dim obj As Variant
    On Error GoTo Err
        ContainsObject = True
        Set obj = col(key)
        Exit Function
Err:

    ContainsObject = False
End Function

' whether an scalar exists in a collection
Function ContainsNonObject(col, key As Variant) As Boolean
    Dim obj As Variant
    On Error GoTo Err
        ContainsNonObject = True
        obj = col(key)
        Exit Function
Err:

    ContainsNonObject = False
End Function
