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


Function TableColumnToArray(TableDicts As Collection, ColumnName As String) As Variant()
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


