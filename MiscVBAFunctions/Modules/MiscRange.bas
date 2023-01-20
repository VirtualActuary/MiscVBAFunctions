Attribute VB_Name = "MiscRange"
Option Explicit

Function ActiveRowsDown(Optional R As Range) As Long
    ' number of active rows down from the starting range
    ' it's capped at a minimum of 1
    '
    ' Args:
    '   R: Input range
    '
    ' Returns:
    '   Number of rows down from the starting range.
    
    If R Is Nothing Then Set R = Selection
    ActiveRowsDown = R.Worksheet.Cells(R.Worksheet.Cells.Rows.Count, R.Column).End(xlUp).Row - R.Row + 1
    ActiveRowsDown = Application.WorksheetFunction.Max(1, ActiveRowsDown)
End Function


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


