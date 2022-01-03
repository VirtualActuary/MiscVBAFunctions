Attribute VB_Name = "MiscTables"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function HasLO(Name As String, Optional WB As Workbook) As Boolean

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If LCase(Name) = LCase(LO.Name) Then
                HasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    HasLO = False

End Function


' get list object only using it's name from within a workbook
Public Function GetLO(Name As String, Optional WB As Workbook) As ListObject

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If LCase(Name) = LCase(LO.Name) Then
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

Function TableToArray( _
      Name As String _
    , Optional WB As Workbook _
    ) As Variant()
    
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


Function GetAllTables(WB As Workbook) As Collection
    Set GetAllTables = New Collection
    
    ' Returns all (potential) tables in a workbook
    
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
    ' Converts a table's column to a 1-dimensional array
    
    Dim arr() As Variant
    ReDim arr(TableDicts.Count - 1) ' zero indexed
    Dim dict As Dictionary
    Dim counter As Long
    For Each dict In TableDicts
        arr(counter) = fn.dictget(dict, ColumnName)
        counter = counter + 1 ' zero indexing
    Next dict
    
    TableColumnToArray = arr
End Function
