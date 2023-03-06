Attribute VB_Name = "MiscRegex"
Option Explicit

Public Function RenameVariableInFormula( _
    InputFormula As String, _
    OldName As String, _
    NewName As String, _
    Optional IgnoreStrings As Boolean = True)
    ' Replaces a variable within an Excel formula
    '
    ' Args:
    '   InputFormula: formula containing the variables to rename
    '   OldName: Current name of the variable
    '   NewName: New name of the variable
    '   IgnoreStrings: Option to ignore matching variables found within a formula string
    ' Returns:
    '   The formula with the variable name replaced
     
    Dim ReString As RegExp
    Set ReString = New RegExp
    With ReString
        ' Regular expression for a quote-escapable Excel string.
        If IgnoreStrings Then
            .Pattern = "\""(?:[^\""]*\""\"")*[^\""]*\""(?!\"")"
        Else
            .Pattern = "^_^"  ' Never match
        End If
        .Global = True
    End With
    
    Dim StringMatches As MatchCollection
    Set StringMatches = ReString.Execute(InputFormula)
     
    Dim Re As RegExp
    Set Re = New RegExp
    With Re
        .Pattern = "\w+"
        .Global = True
    End With
    
    Dim Matches As MatchCollection
    Set Matches = Re.Execute(InputFormula)
    
    Dim Filtered As Collection
    Set Filtered = New Collection
    
    Dim MatchName As Match
    Dim StrMatch As Match
    For Each MatchName In Matches
        If LCase(MatchName.Value) = LCase(OldName) Then
            For Each StrMatch In StringMatches
                If StrMatch.FirstIndex <= MatchName.FirstIndex And MatchName.FirstIndex < StrMatch.FirstIndex + StrMatch.Length Then
                    GoTo ContinueForLoop
                End If
            Next
            Filtered.Add MatchName
        End If
ContinueForLoop:
    Next
    
    RenameVariableInFormula = InputFormula
    Dim I As Long
    For I = Filtered.Count To 1 Step -1
        RenameVariableInFormula = Mid(RenameVariableInFormula, 1, Filtered(I).FirstIndex) & NewName & Mid(RenameVariableInFormula, Filtered(I).FirstIndex + Filtered(I).Length + 1)
    Next
    
End Function

