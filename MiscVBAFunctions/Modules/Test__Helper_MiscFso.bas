Attribute VB_Name = "Test__Helper_MiscFso"
Option Explicit

Function Test_GetAllFilesRecursive(InputPath As String)
    Set Test_GetAllFilesRecursive = New Collection
    
    Dim F As File
    For Each F In GetAllFilesRecursive(Fso.GetFolder(InputPath))
        Test_GetAllFilesRecursive.Add F.Path
    Next F
End Function
