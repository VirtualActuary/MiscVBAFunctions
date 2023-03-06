Attribute VB_Name = "MiscPowerQuery"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

' Helpful functions to help with Power Query manipulations in VBA

Private Sub MiscPowerQueryTests()
    Debug.Print DoesQueryExist("foo"), False
End Sub


Public Function DoesQueryExist(ByVal QueryName As String, Optional WB As Workbook) As Boolean
    ' Check if a Query exists in the given Workbook.
    '
    ' Args:
    '   queryName: Name of the Query to look for.
    '   WB: Name of the WorkBook to look in.
    '
    ' Returns:
    '   True if the Query exists, False otherwise.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Helper function to check if a query with the given name already exists
    Dim Qry As WorkbookQuery
    For Each Qry In WB.Queries
        If (Qry.Name = QueryName) Then
            DoesQueryExist = True
            Exit Function
        End If
    Next
    DoesQueryExist = False
End Function


Public Function GetQuery(Name As String, Optional WB As Workbook) As WorkbookQuery
    ' Return the desired Query if it exists. If the Query doesn't exist, an error is raised.
    '
    ' Args:
    '   Name: Name of the Query to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The desired Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim Qry As WorkbookQuery
    For Each Qry In WB.Queries
        If Qry.Name = Name Then
            Set GetQuery = Qry
            Exit Function
        End If
    Next Qry
    
    Err.Raise 999, , "Query " & Name & " does not exist"
    
End Function

Public Function GetQueryFormula(Name As String, Optional WB As Workbook) As String
    ' Returns the Power Query M formula of a WorkbookQuery
    '
    ' Args:
    '   Name: Name of the Query to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The Power Query M formula of the WorkbookQuery
    GetQueryFormula = GetQuery(Name, WB).Formula
End Function

Public Function UpdateQuery(Name As String, QueryFormula As String, Optional WB As Workbook) As WorkbookQuery
    ' Update the selected Query. If the Query doesn't exist, a new Query is added.
    '
    ' Args:
    '   Name: Name of the Query.
    '   queryFormula: New Formula of the Query.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   Updated or new Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a query to the new formula
    ' if the query doesn't exist, a new one is created
    
    If DoesQueryExist(Name, WB) Then
        Set UpdateQuery = GetQuery(Name, WB)
        UpdateQuery.Formula = QueryFormula
    Else
        Set UpdateQuery = WB.Queries.Add(Name, QueryFormula)
    End If
    
End Function

Public Function UpdateQueryAndRefreshListObject(Name As String, QueryFormula As String, Optional WB As Workbook) As WorkbookQuery
    ' Update the selected Query and refresh the list of objects.
    '
    ' Args:
    '   Name: Name of the Query to update.
    '   queryFormula: New Formula of the Query.
    '   WB: The selected Workbook.
    '
    ' Returns:
    '   Updated or new Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a power query query
    ' Also waits for the query to refresh before continuing the code
    
    ' assumes the ListObject and Query has the same name
    Set UpdateQueryAndRefreshListObject = UpdateQuery(Name, QueryFormula, WB)
    
    WaitForListObjectRefresh Name, WB
    
End Function


Public Sub WaitForListObjectRefresh(Name As String, Optional WB As Workbook)
    ' Refresh elements in the QueryTable.
    '
    ' Args:
    '   Name: Name of the ListObject.
    '   WB: Name of the WorkBook.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Refreshes the query before continuing the code
    
    Dim LO As ListObject
    Set LO = GetLO(Name, WB)
    Dim BGRefresh As Boolean
    With LO.QueryTable
        BGRefresh = .BackgroundQuery
        .BackgroundQuery = False
        .Refresh
        .BackgroundQuery = BGRefresh
    End With
    
End Sub

Public Sub LoadToWorkbook(QueryName As String, Optional WB As Workbook)
    ' loads a query to a sheet in the workbook
    '
    ' Args:
    '   queryName: Name of the query to load to the WorkBook
    '   WB: Name of the WorkBook.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim LO As ListObject
    If HasLO(QueryName, WB) Then
        Set LO = GetLO(QueryName, WB)
        LO.Refresh
    Else
        Dim WS As Worksheet
        Set WS = WB.Worksheets.Add(After:=ActiveSheet)
        WS.Name = NewSheetName(QueryName, ThisWorkbook)
        
        With WS.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = XlCmdSql
            .CommandText = Array("SELECT * FROM [" & QueryName & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = XlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .ListObject.DisplayName = QueryName
            .Refresh BackgroundQuery:=False
        End With
        
    End If
    
End Sub

Public Function AddToWorkbookConnections(Query As WorkbookQuery, Optional WB As Workbook) As WorkbookConnection
    ' adds a query to workbookconnections so that it can be used in pivot tables
    '
    ' Args:
    '   Query: Query that gets added to the workbookconnections.
    '   WB: Name of the WorkBook
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim ConnectionName As String, CommandString As String, CommandText As String, CommandType
    ConnectionName = "Query - " & Query.Name
    CommandString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Query.Name & ";Extended Properties="""""
    CommandText = "SELECT * FROM [" & Query.Name & "]"
    CommandType = 2
    
    ' This code loads the query to the workbook connections
    If HasKey(WB.Connections, ConnectionName) Then
        Set AddToWorkbookConnections = WB.Connections(ConnectionName)
        AddToWorkbookConnections.OLEDBConnection.Connection = CommandString
        AddToWorkbookConnections.OLEDBConnection.CommandText = CommandText
        AddToWorkbookConnections.OLEDBConnection.CommandType = CommandType
    Else
        Set AddToWorkbookConnections = _
        WB.Connections.Add2(ConnectionName, _
            "Connection to the '" & Query.Name & "' query in the workbook.", _
            CommandString _
            , CommandText, CommandType)
        ' should not be loaded to the data model, else we cannot link two pivots to the same cache linking from this query
    End If

End Function



Public Sub RefreshAllQueriesAndPivots(Optional WB As Workbook)
    ' Refresh all Queries and Pivots.
    '
    ' Args:
    '   WB: Name of the WorkBook
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    WB.RefreshAll
End Sub


