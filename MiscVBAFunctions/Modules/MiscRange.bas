Attribute VB_Name = "MiscRange"
Option Explicit

Function RangeToLO(WS As Worksheet, Data As Range, TableName As String) As ListObject
    ' Create a Table containing the data from the input range.
    ' The Range determines the starting Cell of the table.
    '
    ' Args:
    '   WS: Worksheet to add the Table to
    '   Data: Range with the data and cell locations.
    '   TableName: Name of the new table.
    '
    ' Returns:
    '   The Table as a ListObject
    If HasLO(TableName, WS.Parent) = True Then
        Err.Raise FileAlreadyExists, , "Table already exists. "
        Exit Function
    End If

    Set RangeToLO = WS.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=Data, _
        xlListObjectHasHeaders:=xlYes, _
        tablestyleName:="TableStyleMedium2" _
    )
    RangeToLO.Name = TableName
End Function


Function IsInRange(RangeInput As Range, Lookup As Variant) As Boolean
    ' Check if a value in in a Range object.
    '
    ' Args:
    '   RangeInput: Range object to search in
    '   Lookup: Value to search for.
    '
    ' Returns:
    '   True if the value exists in the Range object, False otherwise.
    
    Dim Cell As Range
    Dim Arr() As Variant
    Arr = RangeToArray(RangeInput)
    IsInRange = IsInArray(Arr, Lookup)
End Function


