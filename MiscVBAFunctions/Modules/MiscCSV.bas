Attribute VB_Name = "MiscCSV"
Option Explicit

Function CsvToLO( _
        StartCell As Variant, _
        FilePath As String, _
        TableName As String _
    ) As ListObject
    ' loads a CSV file to an Excel list object
    ' this doesn't support sql_queries like loadTextToLO
    ' we import a query table, refresh and then delete it again
    ' hence, we won't have a live connection to the underlying data.
    ' The Column names are converted to strings.
    '
    ' Args:
    '   StartCell: Starting cell of the new Table
    '   FilePath: Path to the CSV file
    '   TableName: Name of the new Table
    '
    ' Returns:
    '   The new Table as a ListObject
    
    Dim QT As QueryTable
    Set QT = CsvToQueryTable(StartCell.Worksheet, StartCell, FilePath)
    Dim ResultRange As Range
    Set ResultRange = QT.ResultRange
    
    ' we need to delete the query table in order to link it to a list object:
    ' see https://docs.microsoft.com/en-us/office/vba/api/excel.querytable.listobject
    ' data from text query is imported as a QueryTable object, while all other external data is imported as a ListObject object.
    QT.Delete
    
    Set CsvToLO = RangeToLO(StartCell.Worksheet, ResultRange, TableName)
End Function


Private Function CsvToQueryTable(WS As Worksheet, _
                 StartRef, _
                 FilePath As String) As QueryTable
    ' csv file to query table
    ' cannot apply sql query like in loadTextToLO()
    
    Set CsvToQueryTable = WS.QueryTables.Add( _
            Connection:="TEXT;" & FilePath, _
            Destination:=StartRef)
            
    CsvToQueryTable.TextFileCommaDelimiter = True
    ' can set this for queryTables :D
    CsvToQueryTable.TextFileDecimalSeparator = "."
    
    CsvToQueryTable.Refresh
    
End Function
