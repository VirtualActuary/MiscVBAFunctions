Attribute VB_Name = "MiscCreateTextFile"
Option Explicit

Private Sub testCreateTextFile()
    CreateTextFile "foo", ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt"
    ' TODO: assertion
End Sub

Public Function CreateTextFile(ByVal Content As String, ByVal FilePath As String)
    ' Creates a new / overwrites an existing text file with Content
    
    Dim oFile As Integer
    oFile = FreeFile
    
    Open FilePath For Output As #oFile
        Print #oFile, Content
    Close #oFile

End Function
