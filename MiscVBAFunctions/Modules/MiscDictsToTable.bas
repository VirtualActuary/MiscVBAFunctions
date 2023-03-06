Attribute VB_Name = "MiscDictsToTable"
Option Explicit

Public Function DictsToTable( _
        TableDicts As Collection, _
        Start_range As Range, _
        TableName As String, _
        Optional EscapeFormulas As Boolean = False _
    ) As ListObject

    ' Converts a TableToDicts object back to a list object.
    '
    ' Args:
    '   TableDicts: Collection of dictionaries that represents a table [{"col1": 1, "col2": 2}, {"col1": 3, "col2": 4}]
    '   OutputStartRange: 1x1 range where the first header of the table should start
    '   EscapeFormulas: Optional Boolean input.
    '                   If True, formulas get copied as text. (E.x. "=d" -> "'=d")
    '                   If False, the data is copied as is.
    '                   If this is False and "=[foo]" gets copied, the function will crash.
    '
    ' Returns:
    '   ListObject of the new table

    Dim NumberOfRows, NumberOfColumns, I, J As Integer
    Dim CurrentTable As ListObject
    Dim Dict As Dictionary
    Dim ColumnNames() As Variant
    Dim ColumnNamesAsString As String
    Dim DictEntry As Variant
    
    NumberOfRows = TableDicts.Count
    NumberOfColumns = TableDicts(1).Count
    
    ColumnNames = TableDicts(1).Keys()
    ColumnNamesAsString = Join(ColumnNames, ",")

    For Each Dict In TableDicts
        If Dict.Count <> NumberOfColumns Then
            Err.Raise -997, , "Mismatch lengths for the dictionary entries. "
        End If
        
        For Each DictEntry In Dict.Keys()
        
            If (InStr(ColumnNamesAsString, DictEntry) = 0) Then
                Err.Raise -996, , "Mismatching dictionaries found. "
            End If
        Next DictEntry
    Next Dict
    
    Dim Arr() As Variant
    Arr = DictsToArray(TableDicts)
    Set DictsToTable = ArrayToNewTable(TableName, Arr, Start_range, EscapeFormulas)
End Function


Private Sub SetCellValue(Cell As Range, Value As Variant)
    ' Sets a cell's value, and makes allowance for Excel's unwanted
    ' autocorrecting of strings starting with `=`. Using VBA you can set a formula
    ' Explicitly using Cell.Formula = FormulaString instead of relying on Excel's autorcorrect
    '
    ' Args:
    '   Cell: the cell set
    '   Value: the value to give the cell
    '
    ' Returns:
    '   Returns nothing. Side-effect on the value of the cell
    
    If Application.WorksheetFunction.IsText(Value) Then
        If Left(Value, 1) = "=" Then
            With Cell
                .NumberFormat = "@" ' Format as TEXT. It avoids the auto-correction of '=foo' to =foo
                .Value = Value
            End With
            Exit Sub
        End If
    End If
    
    Cell.Value = Value
        
End Sub
