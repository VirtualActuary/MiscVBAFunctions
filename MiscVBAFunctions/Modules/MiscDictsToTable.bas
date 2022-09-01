Attribute VB_Name = "MiscDictsToTable"
Option Explicit

Public Function DictsToTable(TableDicts As Collection, start_range As Range, TableName As String) As ListObject

    ' Converts a TableToDicts object back to a list object.
    '
    ' Args:
    '   TableDicts: Collection of dictionaries that represents a table [{"col1": 1, "col2": 2}, {"col1": 3, "col2": 4}]
    '   OutputStartRange: 1x1 range where the first header of the table should start
    '
    ' Returns:
    '   ListObject of the new table

    Dim NumberOfRows, NumberOfColumns, I, J As Integer
    Dim CurrentTable As ListObject
    Dim dict As Dictionary
    Dim ColumnNames() As Variant
    Dim ColumnNamesAsString As String
    Dim DictEntry As Variant
    
    NumberOfRows = TableDicts.Count
    NumberOfColumns = TableDicts(1).Count
    
    ColumnNames = TableDicts(1).Keys()
    ColumnNamesAsString = Join(ColumnNames, ",")

    For Each dict In TableDicts
        If dict.Count <> NumberOfColumns Then
            Err.Raise -997, , "Mismatch lengths for the dictionary entries. "
        End If
        
        For Each DictEntry In dict.Keys()
        
            If (InStr(ColumnNamesAsString, DictEntry) = 0) Then
                Err.Raise -996, , "Mismatching dictionaries found. "
            End If
        Next DictEntry
    Next dict
    
    ' Add column headers.
    For I = 1 To NumberOfColumns
        start_range.Offset(0, I - 1).Value = TableDicts(1).Keys()(I - 1)
    Next I
    
    ' add values.
    For J = 1 To NumberOfRows
        For I = 1 To NumberOfColumns
            SetCellValue start_range.Offset(J, I - 1), TableDicts(J)(ColumnNames(I - 1))
        Next I
    Next J
    
    Set CurrentTable = start_range.Worksheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=start_range.Resize(NumberOfRows + 1, NumberOfColumns), _
    xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium2")
    
    CurrentTable.Name = TableName

    Set DictsToTable = CurrentTable
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
