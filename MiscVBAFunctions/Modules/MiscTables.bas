Attribute VB_Name = "MiscTables"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function HasLO(Name As String, Optional WB As Workbook) As Boolean
    ' Check if the selected WorkBook contains a ListObject with the input name.
    '
    ' Args:
    '   Name: Name of the ListObject to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   True if the ListObject exists.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If VBA.LCase(Name) = VBA.LCase(LO.Name) Then
                HasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    HasLO = False

End Function


Public Function GetLO(Name As String, Optional WB As Workbook) As ListObject
    ' get list object only using it's name from within a workbook
    '
    ' Args:
    '   Name: Name of the ListObject to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The ListObject if it exists. An error is raised if it doesn't exist.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If VBA.LCase(Name) = VBA.LCase(LO.Name) Then
                Set GetLO = LO
                Exit Function
            End If
        Next LO
    Next WS
    
    If GetLO Is Nothing Then
        ' 9: Subscript out of range
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "List object '" & Name & "' not found in workbook '" & WB.Name & "'")
    End If

End Function


Private Sub TestTableToArray()
    TableToArray "foo"
End Sub

Public Function TableToArray( _
      Name As String _
    , Optional WB As Workbook _
    ) As Variant()
    ' Return an Array of the input table.
    '
    ' Args:
    '   Name: Name of the table to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   2D array of the selected Table.
    
    TableToArray = RangeTo2DArray(TableRange(Name, WB))
    
End Function

Public Function TableRange( _
        Name As String _
      , Optional WB As Workbook _
      ) As Range
    
    'Returns the range (including headers of a table named `Name` in workbook `WB`): _
    - It first looks for a list object called `Name` _
      - If the `.DataBodyRange` property is nothing the table range will only be the headers _
    - Then it looks for a named range in the Workbook scope called `Name` and returns the _
      range this named range is referring to _
    - Then it looks for a worksheet scoped named range called `Name`. The first occurrence _
      will be returned _
    If no tables found, a `SubscriptOutOfRange` error (9) is raised _
    The name of the table to be found is case insensitive
    '
    ' Args:
    '   Name: Name of the table to look for.
    '   WB: Selected Workbook.
    '
    ' Returns:
    '   Range of the cells in the selected Table.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    If HasLO(Name, WB) Then
        Dim LO As ListObject
        Set LO = GetLO(Name, WB)
        If LO.DataBodyRange Is Nothing Then
            Set TableRange = LO.HeaderRowRange
        Else
            Set TableRange = LO.Range
        End If
        Exit Function
    End If
    
    If hasKey(WB.Names, Name) Then
        Set TableRange = WB.Names(Name).RefersToRange
        Exit Function
    End If
    
    Dim WS As Worksheet
    ' this will find the first occurrence of the table called 'Name'
    For Each WS In WB.Worksheets
        If hasKey(WS.Names, Name) Then
            Set TableRange = WS.Names(Name).RefersToRange
            Exit Function
        End If
    Next WS
    
    Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Table '" & Name & "' not found in workbook '" & WB.Name & "'")
    
End Function


Public Function GetAllTables(WB As Workbook) As Collection
    Set GetAllTables = New Collection
    ' Returns all tables in a workbook
    '
    ' Args:
    '   WB: The selected WorkBook
    '
    ' Returns:
    '   All tables in the selected WorkBook.
    
    Dim WS As Worksheet
    Dim LO As ListObject
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            GetAllTables.Add LO.Name
        Next LO
    Next WS
    
    Dim Name As Name
    For Each Name In WB.Names
        GetAllTables.Add Name.Name
    Next Name
    
    For Each WS In WB.Worksheets
        For Each Name In WS.Names
            ' remove the sheetname prefix to get the table name
            GetAllTables.Add Mid(Name.Name, InStr(Name.Name, "!") + 1)
        Next Name
    Next WS
    
End Function


Public Function TableColumnToArray(TableDicts As Collection, ColumnName As String) As Variant()
    ' Append the selected key's value from each Dict in the input Collection to a 1-dimensional array
    '
    ' Args:
    '   TableDicts: A collection of Dicts.
    '   ColumnName: Name of the column that will be returned as a 1-D array.
    '
    ' Returns:
    '   1-D array of the selected column.
    
    Dim arr() As Variant
    ReDim arr(TableDicts.Count - 1) ' zero indexed
    Dim dict As Dictionary
    Dim counter As Long
    For Each dict In TableDicts
        arr(counter) = dictget(dict, ColumnName)
        counter = counter + 1 ' zero indexing
    Next dict
    
    TableColumnToArray = arr
End Function


Public Sub CopyTable(InputTableName As String _
                    , StartRange As Range _
                    , Optional OutputTableName As String _
                    , Optional InputWB As Workbook)
    ' Copy a List Object or TableRange to the desired location as a Table.
    ' This can be in the same Workbook, or a different workbook.
    ' The of the output Table NumberFormat is the same as the input table's.
    '
    ' Args:
    '   InputTableName: Table name that will be copied.
    '   StartRange: Range object of the output table's destination
    '   OutputTableName: Name of the output table.
    '                    If left empty, the input table's name gets used.
    '   InputWB: WorkBook of the input Table. ThisWorkBook is used if left empty.

    Dim col1 As Collection
    Dim InputTableRange As Range
    Dim OutputTableRange As Range
    Dim I As Long
    
    If OutputTableName = "" Then
        OutputTableName = InputTableName
    End If
    If InputWB Is Nothing Then Set InputWB = ThisWorkbook

    Set col1 = TableToDicts(InputTableName, InputWB)
    Set InputTableRange = TableRange(InputTableName, InputWB)
    Set OutputTableRange = DictsToTable(col1, StartRange, OutputTableName).Range

    For I = 1 To InputTableRange.Count
        OutputTableRange(I).NumberFormat = InputTableRange(I).NumberFormat
    Next
    
End Sub


Function TableColumnToCollection(TableDicts As Collection, ColumnName As String) As Collection
    ' Append the selected key's value from each Dict in the input Collection to a Collection
    '
    ' Args:
    '   TableDicts: A collection of Dicts.
    '   ColumnName: Name of the column that will be returned as a Collection.
    '
    ' Returns:
    '   Collection of the selected column.
    
    Dim col1 As Collection
    Dim dict As Dictionary
    
    Set col1 = New Collection
    For Each dict In TableDicts
        col1.Add dictget(dict, ColumnName)
    Next dict
    
    Set TableColumnToCollection = col1
End Function


Public Sub ResizeLO(LO As ListObject, NumRows As Long)
    ' Resize a table to the desired number of data rows. The columns remain unchanged.
    ' If NumRows is set to "0", the table will instead be resized to "1" and the
    ' content in that 1st row will be cleared.
    '
    ' Args:
    '   LO: Selected List Object.
    '   NumRows: Number of desired rows.
    
    If NumRows = 0 Then
        LO.DataBodyRange.Delete
        Exit Sub
    End If
    
    Dim oldNumRows As Long
    oldNumRows = LO.ListRows.Count
    
    ' Resize the table (add 1 to the number of rows because mListObject.Range
    ' includes the header row).
    LO.Resize _
        LO.Range.Resize( _
            NumRows + 1, _
            LO.ListColumns.Count)
    
    ' If the table is resized to have one row, but the row contains no data,
    ' the row will be treated as the Insert row, and the data row count will
    ' remain zero.  This will cause problems since the table doesn't actually
    ' have a DataBodyRange.  To avoid this situation, put a space in the first
    ' column, which will cause the Insert row to change to a data row.  After
    ' setting the value once, it can be removed and the row will remain part
    ' of the DataBodyRange.
    
    If NumRows = 1 And LO.ListRows.Count = 0 Then
        LO.ListRows.Add
    End If
    
    ' If the new number of rows is less than the old number of rows, clear out
    ' the rows that were just removed from the table.
    If NumRows < oldNumRows Then
        LO.DataBodyRange _
            .Offset(NumRows, 0) _
            .Resize(oldNumRows - NumRows, LO.ListColumns.Count) _
            .ClearContents
    End If
End Sub


Public Function GetTableColumnRange( _
      TableName As String _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    ' Returns the range of a table's column, including the header
    '
    ' Args:
    '   TableName: Name of the Table.
    '   Columns: Name of the column.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   Range of cells for the selected table's column.
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    Dim I As Long
    For I = 1 To TableR.Columns.Count
        If VBA.LCase(TableR(1, I).Value) = VBA.LCase(Column) Then
            GoTo found
        End If
    Next I
    
    Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Column '" & Column & "' not found in table '" & TableName & "'")
found:
    ' Intersect of table range and entirecolumn
    Set GetTableColumnRange = Intersect(TableR, TableR(1, I).EntireColumn)

End Function


Public Function GetTableColumnDataRange(LO As ListObject, ColumnName As String) As Range
    ' Returns the data range for the given column of this Excel table.
    ' An error will be raised if the selected Column name doesn't exist in the given List Object.
    '
    ' Args:
    '   LO: List Object to fetch the column from
    '   ColumnName: Name of the column where the data will be fetched from.
    '
    ' Returns:
    '   Data Range of the given column.
    
    Dim listCol As ListColumn
    
    If Not hasKey(LO.ListColumns, ColumnName) Then
        Err.Raise 32000, Description:= _
        "Column: '" & ColumnName & "' of table '" _
            & LO.Name & "' Doesn't exist."
    End If
    
    Set listCol = LO.ListColumns(ColumnName)
    
    'On Error GoTo noDataRange
    If LO.DataBodyRange Is Nothing Then
        Exit Function
    End If
    
    Set GetTableColumnDataRange = listCol.DataBodyRange
'    Exit Function
'
'noDataRange:
'    On Error GoTo 0
    
End Function


Public Function GetTableRowRange( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Range
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, a runtime error is raised
    '
    ' Args:
    '   TableName: Name of the Table
    '   Columns: Collection of Column names.
    '   Values: Values to match agains.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The row in which the vales matches the comparison.
    
    Dim RowNumber As Long
    RowNumber = GetTableRowIndex(TableName, Columns, Values, WB) ' this will throw a runtime error if not found
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    ' Intersect of table range and entirerow
    ' +1 as header is not included in GetTableRowIndex
    Set GetTableRowRange = Intersect(TableR, TableR(RowNumber + 1, 1).EntireRow)
End Function


Public Sub GotoRowInTable( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    )
    ' Go to the cell that matches the entry in the Values input.
    '
    ' Args:
    '   TableName: Name of the Table.
    '   Columns: Columns to include in the search.
    '   Values: Values to search for.
    '   WB: Selected WorkBook.
    
    Application.GoTo GetTableRowRange(TableName, Columns, Values, WB), True
End Sub


Public Function GetTableRowNumberDataRange(LO As ListObject, RowNumber As Long) As Range
    ' Returns the data range for the given row number of this Excel table.
    ' An error will be raised if the selected row number name doesn't exist in the given List Object.
    '
    ' Args:
    '   LO: List Object to fetch the column from
    '   RowNumber: Row number where the data will be fetched from.
    '
    ' Returns:
    '   Data Range of the given row number.
    
    Dim listCol As ListColumn
    Dim listR As listRow
        
    If Not hasKey(LO.ListRows, RowNumber) Then
        Err.Raise 32000, Description:= _
        "Row number '" & RowNumber & "' of table '" _
            & LO.Name & "' doesn't exist."
    End If
    
    Set listR = LO.ListRows(RowNumber)
    Set GetTableRowNumberDataRange = listR.Range
End Function
