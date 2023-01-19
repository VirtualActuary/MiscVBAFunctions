Attribute VB_Name = "MiscTextFile"
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
    
    Dim OFile As Integer
    OFile = FreeFile
    
    Open FilePath For Output As #OFile
        Print #OFile, Content
    Close #OFile

End Sub


Function readTextFile(Path As String) As String
    Dim OFile As Object
    Set OFile = Fso.OpenTextFile(Path, ForReading)
    
    readTextFile = OFile.ReadAll
    
End Function

