Attribute VB_Name = "MiscTables"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function HasLO(Name As String, Optional WB As Workbook) As Boolean

    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Dim WS As Worksheet, LO As ListObject
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
        Err.Raise 9, , "List object '" & Name & "' not found in workbook '" & WB.Name & "'"
    End If

End Function
