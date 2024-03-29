Attribute VB_Name = "MiscTableToDicts"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Private Sub TableToDictsTest()
    Dim Dicts As Collection
    Set Dicts = TableToDicts("TableToDictsTestData")
    ' read row 2 in column "b":
    Debug.Print Dicts(2)("b"), 5
End Sub

Public Function TableToDictsLogSource( _
          TableName As String _
        , Optional WB As Workbook _
        , Optional Columns As Collection _
        ) As Collection
    
    'Similar to TableToDicts, but also stores the source of each row _
    in a dictionary with key `__source__`
    'The `__source__` object contains the following keys: _
     - `Workbook`: the Workbook object with the table _
     - `Table`: the name of the table within the workbook _
     - `RowIndex`: the row index of the current entry of the table
    '
    ' Args:
    '   TableName: Name of the table to convert to Dicts.
    '   WB: Selected WorkBook
    '   Columns: Columns to include in the Dicts.
    '
    ' Returns:
    '   The collection of Dicts containing the info as well as the source of each row.
    
    Set TableToDictsLogSource = TableToDicts(TableName, WB, Columns)
    Dim Dict As Dictionary
    Dim RowIndex As Long
    RowIndex = 0
    For Each Dict In TableToDictsLogSource
        RowIndex = RowIndex + 1
        Dict.Add "__source__", DictI("Workbook", WB, "Table", TableName, "RowIndex", RowIndex)
    Next Dict
End Function


Public Function TableToDicts( _
          TableName As String _
        , Optional WB As Workbook _
        , Optional Columns As Collection _
        ) As Collection
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/8e1b9709a380d71eaf0d59bd0c2882c8501e9540/aa_py_xl/data_util.py#L21
    ' Convert a Table to a Collection of Dicts.
    ' Column names are case insensitive, i.e. `c` and `C` will be treated as duplicate column names.
    ' When columns are duplicated the last instance of the column name is used.
    '
    ' Args:
    '   TableName: Name of the Selected Table.
    '   WB: Selected WorkBook
    '   Columns: Columns to be added to the Dicts.
    '
    ' Returns:
    '   A collection of Dictionaries.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set TableToDicts = New Collection
    
    Dim D As Dictionary
    
    Dim I As Long
    Dim J As Long
    Dim TableData() As Variant
    TableData = TableToArray(TableName, WB)
    
    For I = LBound(TableData, 1) + 1 To UBound(TableData, 1)
        Set D = New Dictionary
        D.CompareMode = TextCompare ' columns are case insensitive
        
        If Columns Is Nothing Then
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                D(TableData(0, J)) = TableData(I, J)
            Next J
        Else
            Dim ColumnName As Variant
            Dim Column As Variant
            
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                ColumnName = TableData(LBound(TableData, 2), J)
                If IsValueInCollection(Columns, ColumnName) Then
                    D(ColumnName) = TableData(I, J)
                End If
            Next J
        End If
        
        TableToDicts.Add D
    Next I
    
End Function

Private Function TestGetTableRowIndex()
    Dim Table As Collection
    Set Table = Col(DictI("a", 1, "b", 2), DictI("a", 3, "b", 4), DictI("a", "foo", "b", "bar"))
    Debug.Print GetTableRowIndex(Table, Col("a", "b"), Col(3, 4)), 2
    Debug.Print GetTableRowIndex(Table, Col("a", "b"), Col("foo", "bar")), 3
End Function


Public Function TableLookupValue( _
        Table As Variant _
      , Columns As Collection _
      , Values As Collection _
      , ValueColName As String _
      , Optional Default As Variant = Empty _
      , Optional WB As Workbook _
      ) As Variant
    ' Returns the value from the ValueColName column in a TableToDicts object _
      given the value In the lookup column _
      A default value can be assigned For when no lookup Is found _
      Otherwise it returns a runtime Error
    '
    ' Args:
    '   Table: Selected table.
    '   Columns: Collection of selected Column names.
    '   Values: Values from the lookup column
    '   ValueColName: Column name that gets used to fetch values from.
    '   default: Value to be used when no value has been found.
    '   WB: Selected workbook.
    '
    ' Returns:
    '   Value from the ValueColName column.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' for when GetTableRowIndex fails
    If Not IsEmpty(Default) Then On Error GoTo SetDefault
    
    Dim Dict As Dictionary
    Set Dict = EnsureTableDicts(Table, WB)(GetTableRowIndex(Table, Columns, Values, WB))
    TableLookupValue = Dictget(Dict, ValueColName, Default)
    
    Exit Function
SetDefault:
    TableLookupValue = Default
    
End Function


Public Function TableLookupCell( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    ' Find a cell in a Table and return its range.
    ' The first match is returned.
    '
    ' Args:
    '   TableName: Name of the table.
    '   Columns: Columns to use to search the Values
    '   Values: The values to search for.
    '   Column: Name of any column in the table. Is used to determine the size of the table.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   The range of the cell that matches its matching Value first.
    
    Set TableLookupCell = Intersect(GetTableRowRange(TableName, Columns, Values, WB), GetTableColumnRange(TableName, Column, WB))

End Function


Private Function EnsureTableDicts(Table As Variant, Optional WB As Workbook) As Collection
    
    If TypeOf Table Is Collection Then ' assume if collection, it's already a TableDicts object
        Set EnsureTableDicts = Table
    Else
        Set EnsureTableDicts = TableToDicts(CStr(Table), WB)
    End If

End Function


Public Function GetTableRowIndex( _
      Table As Variant _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    , Optional IgnoreCaseValues As Boolean = True _
    ) As Long
    ' Given a table name, Columns and Values to match this function returns the row in which the first set of values matches
    ' Comparison is case sensitive
    ' If no match is found, SubscriptOutOfRange error is raised
    '
    ' Args:
    '   Table: TableToDicts or name of the table to find.
    '   Columns: Columns to match
    '   Values: Values to match.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The row in which the values matches the comparison.
    
    Dim Dict As Dictionary
    Dim KeyValuePair As Collection
    Dim IsMatch As Boolean
    Dim RowNumber As Long
    Dim ValLhs As Variant
    Dim ValRhs As Variant
    
    
    For Each Dict In EnsureTableDicts(Table, WB) ' Already a Dicti
        IsMatch = True
        RowNumber = RowNumber + 1
        For Each KeyValuePair In Zip(Columns, Values)
            Assign ValLhs, Dict(KeyValuePair(1))  ' Allow entries to be objects
            Assign ValRhs, KeyValuePair(2)
            
            If IgnoreCaseValues Then
                If IsString(ValLhs) Then ValLhs = LCase(ValLhs)
                If IsString(ValRhs) Then ValRhs = LCase(ValRhs)
            End If
                
            If ValLhs <> ValRhs Then
                IsMatch = False
            End If
        Next KeyValuePair
        If IsMatch = True Then Exit For
    Next Dict
    
    If IsMatch Then
        GetTableRowIndex = RowNumber
    Else
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Columns-values pairs did not find a match")
    End If
    
End Function


Public Function TableDictToArray(TableDicts As Collection) As Variant()
    ' Convert a TableDicts to an Array. The Column names of the TableDicts
    ' get inserted as the first row in the array.
    '
    ' Args:
    '   TableDicts: Collection of dictionaries as a TableDicts.
    '
    ' Returns:
    '   Array containing the input data.
    
    Dim NumberOfRows As Long
    Dim NumberOfColumns As Long
    Dim I As Integer
    Dim J As Integer
    Dim Dict As Dictionary
    Dim ColumnNames() As Variant
    Dim ColumnNamesAsString As String
    Dim DictEntry As Variant
    
    NumberOfRows = TableDicts.Count
    NumberOfColumns = TableDicts(1).Count
    Dim Arr() As Variant
    ReDim Arr(NumberOfRows, NumberOfColumns - 1)
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

    For I = 0 To UBound(ColumnNames)
        Arr(0, I) = ColumnNames(I)
    Next
    
    For I = 0 To NumberOfRows - 1
        For J = 0 To NumberOfColumns - 1
            Arr(I + 1, J) = TableDicts(I + 1)(ColumnNames(J))
        Next J
    Next I
    TableDictToArray = Arr
End Function


