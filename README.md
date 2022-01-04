# MiscVBAFunctions
 A collection of standalone VBA functions to easily use in other projects.
 
## MiscPowerQuery

### `Public Function doesQueryExist(ByVal queryName As String, Optional WB As Workbook) As Boolean`

Determines whether a query exists in a given workbook.
- `QueryName`: The name of the query to look for
- `WB` [Optional]: The workbook in which to look for the query. Defaults to `ThisWorkbook`


## MiscTables

### TableRange

```
Public Function TableRange( _
        Name As String _
      , Optional WB As Workbook _
      ) As Range
```
Returns the range (including headers of a table named `Name` in workbook `WB`):
- It first looks for a list object called `Name`
  - If the `.DataBodyRange` property is nothing the table range will only be the headers
- Then it looks for a named range in the Workbook scope called `Name` and returns the 
  range this named range is referring to
- Then it looks for a worksheet scoped named range called `Name`. The first occurrence 
  will be returned
If no tables found, a `SubscriptOutOfRange` error (9) is raised
The name of the table to be found is case insensitive

## TableToArray

```
Function TableToArray( _
      Name As String _
    , Optional WB As Workbook _
    ) As Variant()
```
Returns a 2-dimensional array of the contents and headers of a table 
named `Name` in workbook `WB`


## MiscTableToDicts

### TableToDicts

```
Public Function TableToDicts(TableName As String, _
        Optional WB As Workbook, _
        Optional Columns As Collection) As Collection
```

Converts a table to a collection of dictionaries:
- `TableName`: name of the excel Table/Listobject or 
Named range (worksheet level named ranges not implemented)
- `WB` [Optional]: The workbook in which to look for the table
- `Columns`: Collection of the columns to include in the dictionaries

Behaviour:
 - Empty tables will return an initiated collection with zero entries
 - Dictionaries are case insensitive (as ListObject columns are treated in Excel)

### TableToDictsLogSource

Similar to TableToDicts, but also stores the source of each row 
in a dictionary with key `__source__`

The `__source__` object contains the following keys:
 - `Workbook`: the Workbook object with the table
 - `Table`: the name of the table within the workbook
 - `RowIndex`: the row index of the current entry of the table


### GetTableRowIndex

```
Function GetTableRowIndex( _
      Table As Variant _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Long
```

Table can either be a TableToDicts collection, or the name of 
the table to find. 

Given a `Table`, it returns the index of the first row
where all the `Values` matches the values in the corresponding
`Columns`. The matching is case sensitive for strings.

If no match is found a SubscriptOutOfRange error (9) is raised.

### GetTableRowRange

```
Function GetTableRowRange( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Range
```


Given a `Table`, it returns the range of the first row  
where all the `Values` matches the values in the corresponding
`Columns`. The matching is case sensitive for strings.

It returns a SubscriptOutOfRange runtime error (9) if no matches
are found.

### GetTableColumnRange
```
Function GetTableColumnRange( _
      TableName As String _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
```
Returns the range of a table's column, including the header.
Matching is case insensitive.

### TableLookupCell
```
Public Function TableLookupCell( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
```

The intersect of `GetTableRowRange` and `GetTableColumnRange`

### GotoRowInTable

```
Public Sub GotoRowInTable( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    )
```
Goes to the range returned by `GetTableRowRange`

### TableLookupValue

```
Function TableLookupValue( _
        Table As Variant _
      , Columns As Collection _
      , Values As Collection _
      , ValueColName As String _
      , Optional Default As Variant = Empty _
      , Optional WB As Workbook _
      ) As Variant
```

Given a `Table`, it returns the value from the `ValueColName` from the 
first row where all the `Values` matches the values in the corresponding
`Columns`. The matching for `Values` is case sensitive. Matching for column
names are case insensitive.

A default value can be assigned For when no lookup Is found. Otherwise
it returns a SubscriptOutOfRange runtime error (9)

Arguments:
- `Table`: either a TablesToDicts object or the name of a table
- `Columns`: the columns in which to look for the matching `Values`
- `Values`: the values to look for in the given `Column`
- `ValueColName`: the column from which to return the value
- `Default`: 