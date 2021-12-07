# MiscVBAFunctions
 A collection of standalone VBA functions to easily use in other projects.
 
## MiscPowerQuery

### `Public Function doesQueryExist(ByVal queryName As String, Optional WB As Workbook) As Boolean`

Determines whether a query exists in a given workbook.
- `QueryName`: The name of the query to look for
- `WB` [Optional]: The workbook in which to look for the query. Defaults to `ThisWorkbook`