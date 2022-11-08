Attribute VB_Name = "MiscRegex"
Option Explicit

Public Function RenameVariableInFormula( _
    InputFormula As String, _
    OldName As String, _
    NewName As String)
    ' Replaces a variable within a formula
    ' It would also replace strings within quotation marks
    '
    ' Args:
    '   InputFormula: formula containing the variables to rename
    '   OldName: Current name of the variable
    '   NewName: New name of the variable
    '
    ' Returns:
    '   The formula with the variable name replaced
     
    Dim Re As RegExp
    Set Re = New RegExp
    With Re
        .Pattern = "\w+"
        .Global = True
    End With
    
    Dim Matches As MatchCollection
    Set Matches = Re.Execute(InputFormula)
    
    Dim Filtered As Collection
    Set Filtered = col()
    
    Dim Match As Match
    For Each Match In Matches
        If LCase(Match.Value) = LCase(OldName) Then
            Filtered.Add Match
        End If
    Next
    
    RenameVariableInFormula = InputFormula
    Dim I As Long
    For I = Filtered.Count To 1 Step -1
        RenameVariableInFormula = Mid(RenameVariableInFormula, 1, Filtered(I).FirstIndex) & NewName & Mid(RenameVariableInFormula, Filtered(I).FirstIndex + Filtered(I).length + 1)
    Next
    
End Function

