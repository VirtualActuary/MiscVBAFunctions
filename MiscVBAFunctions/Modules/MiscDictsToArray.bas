Attribute VB_Name = "MiscDictsToArray"
Option Explicit


Public Function DictsToArray(TableDicts As Collection) As Variant()
    ' Copies the input TableDicts to a 2D array. The first row in the array
    ' is the keys if the Dicts. the other rows are the copied content.
    '
    ' Args:
    '   TableDicts: Collection of Dicts that will be used to create the array.
    '
    ' Returns:
    '   Array with the copied content
    
    Dim CurrentDict As Dictionary
    Dim DictEntry As Variant
    Dim ColCounter As Long
    Dim DictCounter As Long
    Dim DictLen As Long
    DictLen = TableDicts(1).Count

    Dim Arr() As Variant
    ReDim Preserve Arr(0 To TableDicts.Count, 0 To DictLen - 1)   ' +1 for header
    Dim DictItems() As Variant
    
    Dim Keys() As Variant
    Keys = TableDicts(1).Keys()
    
    For DictCounter = 0 To DictLen - 1
        Arr(0, DictCounter) = TableDicts(1).Keys()(DictCounter)
    Next DictCounter
   
    For ColCounter = 1 To TableDicts.Count
        DictItems = TableDicts(ColCounter).Items()
        For DictCounter = 0 To DictLen - 1
            Arr(ColCounter, DictCounter) = TableDicts(ColCounter)(Keys(DictCounter))
        Next
    Next
    DictsToArray = Arr
End Function
