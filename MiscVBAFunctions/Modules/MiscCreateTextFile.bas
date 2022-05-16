Attribute VB_Name = "MiscCreateTextFile"
Option Explicit

Private Sub testCreateTextFile()
    CreateTextFile "foo", ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt"
    ' TODO: assertion
End Sub

Public Sub CreateTextFile(ByVal Content As String, ByVal FilePath As String)
    ' Creates a new / overwrites an existing text file with Content
    '
    ' Args:
    '   Content: Content that must be inserted into the file.
    '   FilePath: Path where the file will be created. The filename and extension must be included here.
    
    Dim oFile As Integer
    oFile = FreeFile
    
    Open FilePath For Output As #oFile
        Print #oFile, Content
    Close #oFile

End Sub
