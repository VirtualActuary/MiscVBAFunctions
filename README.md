# MiscVBAFunctions
 A collection of standalone VBA functions to easily use in other projects.
 
## MiscPowerQuery

### `Public Function doesQueryExist(ByVal queryName As String, Optional WB As Workbook) As Boolean`

Determines whether a query exists in a given workbook.
- `QueryName`: The name of the query to look for
- `WB` [Optional]: The workbook in which to look for the query. Defaults to `ThisWorkbook`


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