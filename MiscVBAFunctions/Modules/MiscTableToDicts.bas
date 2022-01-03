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

    Set TableToDictsLogSource = TableToDicts(TableName, WB, Columns)
    Dim dict As Dictionary
    Dim RowIndex As Long
    RowIndex = 0
    For Each dict In TableToDictsLogSource
        RowIndex = RowIndex + 1
        dict.Add "__source__", dicti("Workbook", WB, "Table", TableName, "RowIndex", RowIndex)
    Next dict
End Function


Public Function TableToDicts( _
          TableName As String _
        , Optional WB As Workbook _
        , Optional Columns As Collection _
        ) As Collection
    
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/8e1b9709a380d71eaf0d59bd0c2882c8501e9540/aa_py_xl/data_util.py#L21
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set TableToDicts = New Collection
    
    Dim d As Dictionary
    
    Dim I As Long
    Dim J As Long
    Dim TableData() As Variant
    TableData = TableToArray(TableName, WB)
    
    For I = LBound(TableData, 1) + 1 To UBound(TableData, 1)
        Set d = New Dictionary
        d.CompareMode = TextCompare ' must be case insensitive
        
        If Columns Is Nothing Then
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                d.Add TableData(1, J), TableData(I, J)
            Next J
        Else
            Dim ColumnName As Variant
            Dim Column As Variant
            
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                ColumnName = TableData(LBound(TableData, 2), J)
                If IsValueInCollection(Columns, ColumnName) Then
                    d.Add ColumnName, TableData(I, J)
                End If
            Next J
        End If
        
        TableToDicts.Add d
    Next I
    
End Function

Private Function TestGetTableRowIndex()
    Dim Table As Collection
    Set Table = col(dicti("a", 1, "b", 2), dicti("a", 3, "b", 4), dicti("a", "foo", "b", "bar"))
    Debug.Print GetTableRowIndex(Table, col("a", "b"), col(3, 4)), 2
    Debug.Print GetTableRowIndex(Table, col("a", "b"), col("foo", "bar")), 3
End Function


Function TableLookupValue( _
        Table As Variant _
      , Columns As Collection _
      , Values As Collection _
      , ValueColName As String _
      , Optional default As Variant = Empty _
      , Optional WB As Workbook _
      ) As Variant
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' Returns the value from the ValueColName column in a TableToDicts object _
      given the value In the lookup column _
      A default value can be assigned For when no lookup Is found _
      Otherwise it returns a runtime Error
    
    Dim dict As Dictionary
    Set dict = EnsureTableDicts(Table, WB)(GetTableRowIndex(Table, Columns, Values, WB))
    TableLookupValue = dictget(dict, ValueColName, default)

End Function

Function GetTableRowRange( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Range
    
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, a runtime error is raised
    
    Dim RowNumber As Long
    RowNumber = GetTableRowIndex(TableName, Columns, Values, WB) ' this will throw a runtime error if not found
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    ' Intersect of table range and entirerow
    ' +1 as header is not included in GetTableRowIndex
    Set GetTableRowRange = Intersect(TableR, TableR(RowNumber + 1, 1).EntireRow)
    
End Function


Function GetTableColumnRange( _
      TableName As String _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    
' Returns the range of a table's column, including the header
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    Dim I As Long
    For I = 1 To TableR.Columns.Count
        If LCase(TableR(1, I).Value) = LCase(Column) Then
            GoTo found
        End If
    Next J
    
    Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Column '" & Column & "' not found in table '" & TableName & "'")
found:
    ' Intersect of table range and entirecolumn
    Set GetTableColumnRange = Intersect(TableR, TableR(1, I).EntireColumn)

End Function


Public Function TableLookupCell( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    
    Set TableLookupCell = Intersect(GetTableRowRange(TableName, Columns, Values, WB), GetTableColumnRange(TableName, Column, WB))

End Function

Private Function EnsureTableDicts(Table As Variant, Optional WB As Workbook) As Collection
    
    If TypeOf Table Is Collection Then ' assume if collection, it's already a TableDicts object
        Set EnsureTableDicts = Table
    Else
        Set EnsureTableDicts = TableToDicts(CStr(Table), WB)
    End If

End Function


Function GetTableRowIndex( _
      Table As Variant _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Long
    
    ' Table can either be a TableToDicts collection, _
      or the name of the table to find
    
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, SubscriptOutOfRange error is raised
    
    Dim dict As Dictionary
    Dim keyValuePair As Collection
    Dim isMatch As Boolean
    Dim RowNumber As Long
    
    For Each dict In EnsureTableDicts(Table, WB)
        isMatch = True
        RowNumber = RowNumber + 1
        For Each keyValuePair In zip(Columns, Values)
            If dict(keyValuePair(1)) <> keyValuePair(2) Then
                isMatch = False
            End If
        Next keyValuePair
        If isMatch = True Then Exit For
    Next dict
    
    If isMatch Then
        GetTableRowIndex = RowNumber
    Else
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, ":")
    End If
    
End Function

Public Sub GotoRowInTable( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    )
    Application.Goto GetTableRowRange(TableName, Columns, Values, WB), True
End Sub
