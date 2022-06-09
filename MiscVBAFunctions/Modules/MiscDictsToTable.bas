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

    start_range.Worksheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=start_range, _
    xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium2").Name = TableName

    Set CurrentTable = start_range.Worksheet.ListObjects(TableName)
    CurrentTable.ListColumns.Item(1).Name = TableDicts(1).Keys()(0)
    
    ' Add column headers.
    For I = 1 To NumberOfColumns - 1
        CurrentTable.ListColumns.Add.Name = TableDicts(1).Keys()(I)
    Next I
    
    ' add values.
    For J = 1 To NumberOfRows
        CurrentTable.ListRows.Add
        For I = 1 To NumberOfColumns
            CurrentTable.ListRows.Item(J).Range(I) = TableDicts(J)(ColumnNames(I - 1))
        Next I
    Next J
    
    Set DictsToTable = CurrentTable
End Function


