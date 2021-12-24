Attribute VB_Name = "Module1"
Option Explicit

Private Sub testBubbleSort()
    Dim coll As Collection
    Set coll = fn.col("variables_10", "variables", "variables_2")
    Set coll = bubbleSort(coll)
    ' this will currently put {table}_10 before {table}_2
    ' need to fix this for implementations > {table}_9
    Debug.Print coll(1), "variables"
    Debug.Print coll(2), "variables_2" ' :/
    Debug.Print coll(3), "variables_10" ' :/
    
End Sub

Public Function bubbleSort(coll As Collection) As Collection
    
    ' from: https://github.com/austinleedavis/VBA-utilities/blob/f23f1096d8df0dfdc740e5a3bec36525d61a3ffc/Collections.bas#L73
    ' this is an easy implementation but a slow sorting algorithm
    ' do not use for large collections
    
    Dim sortedColl As Collection
    Set sortedColl = New Collection
    Dim vItm As Variant
    ' copy the collection"
    For Each vItm In coll
        sortedColl.Add vItm
    Next vItm

    Dim I As Long, J As Long
    Dim vTemp As Variant

    'Two loops to bubble sort
    For I = 1 To sortedColl.Count - 1
        For J = I + 1 To sortedColl.Count
            If sortedColl(I) > sortedColl(J) Then
                'store the lesser item
               vTemp = sortedColl(J)
                'remove the lesser item
               sortedColl.Remove J
                're-add the lesser item before the
               'greater Item
               sortedColl.Add vTemp, vTemp, I
            End If
        Next J
    Next I
    
    Set bubbleSort = sortedColl
    
End Function

Function getMatchingTables(baseName As String, WB As Workbook) As Collection
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set getMatchingTables = New Collection
    Dim TableName As Variant
    For Each TableName In getAllTables(WB)
        If matchBaseNameAndUnderscoreNumeric(CStr(TableName), baseName) Then
            getMatchingTables.Add TableName
        End If
    Next TableName
    
End Function

Private Sub testMatchBaseName()
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables", "Variables"), True
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables_", "Variables"), False
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables_1", "Variables"), True
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables_2", "Variables"), True
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables_100", "Variables"), True
    Debug.Print matchBaseNameAndUnderscoreNumeric("variables_100e", "Variables"), False
End Sub

Function matchBaseNameAndUnderscoreNumeric(Name As String, baseName As String) As Boolean
    
    matchBaseNameAndUnderscoreNumeric = LCase(Name) = LCase(baseName) Or _
           (startsWith(LCase(Name), LCase(baseName)) And _
                Mid(Name, Len(baseName) + 1, 1) = "_" And _
                IsNumeric(Mid(Name, Len(baseName) + 2)))
End Function


Function getAllTables(WB As Workbook) As Collection
    Set getAllTables = New Collection
    
    Dim WS As Worksheet
    Dim LO As ListObject
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            getAllTables.Add LO.Name
        Next LO
    Next WS
    
    Dim Name As Name
    For Each Name In WB.Names
        getAllTables.Add Name.Name
    Next Name
    
End Function


Function tableColumnToArray(TableDicts As Collection, ColumnName As String) As Variant()
    ' Converts a table's column to a 1-dimensional array
    
    Dim arr() As Variant
    ReDim arr(TableDicts.Count - 1) ' zero indexed
    Dim dict As Dictionary
    Dim counter As Long
    For Each dict In TableDicts
        arr(counter) = fn.dictget(dict, ColumnName)
        counter = counter + 1 ' zero indexing
    Next dict
    
    tableColumnToArray = arr
End Function

Sub allocateTableColumnToArray(TableDicts As Collection, ColumnName As String, arrToAllocate As Variant)
    Dim arr() As Variant
    arr = tableColumnToArray(TableDicts, ColumnName)
    ReDim arrToAllocate(UBound(arr))
    Dim I As Long
    For I = LBound(arr) To UBound(arr)
        arrToAllocate(I) = arr(I)
    Next I
End Sub

