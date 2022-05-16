Attribute VB_Name = "MiscOs"
Option Explicit

Public Function Path(ParamArray entries() As Variant) As String
    ' Combines folder pathes and the name of folders or a file and
    ' returns the combination with valid path separators.
    '
    ' Args:
    '   entries: The folder pathes and the name of folders or a file to be combined.
    '
    ' Returns:
    ' The combination of paths with valid path separators.
    
    Dim Entry As Variant
    
    Path = entries(0)

    For Entry = LBound(entries) + 1 To UBound(entries)
        Path = fso.BuildPath(Path, entries(Entry))
    Next
    
End Function


