Attribute VB_Name = "MiscTableToDicts"
Option Explicit

Private Sub TableToDictsTest()
    Dim Dicts As Collection
    Set Dicts = TableToDicts("TableToDictsTestData")
    ' read row 2 in column "b":
    Debug.Print Dicts(2)("b"), 5
End Sub

Public Function TableToDicts(TableName As String, Optional WB As Workbook) As Collection
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set TableToDicts = New Collection
    
    Dim d As Dictionary
    
    Dim Table As ListObject
    Dim lr As ListRow
    Dim lc As ListColumn
    Set Table = GetLO(TableName, WB)
    For Each lr In Table.ListRows
        Set d = New Dictionary
        For Each lc In Table.ListColumns
            d.Add lc.Name, lr.Range(1, lc.Index).Value
        Next lc
        
        TableToDicts.Add d
    Next lr
    
End Function
