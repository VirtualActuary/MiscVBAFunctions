Attribute VB_Name = "Test__Helper_MiscFso"
Option Explicit

Function Test_GetAllFilesRecursive(InputPath As String)
    Dim AllFiles As Collection
    Dim Pass As Boolean
    Pass = True
    
    Set AllFiles = GetAllFilesRecursive(fso.GetFolder(InputPath))

    Pass = 5 = CInt(AllFiles.Count) = Pass
    Pass = "E:\AA\MiscVBAFunctions\test_data\GetAllFiles\empty file.txt" = AllFiles(1) = Pass
    Pass = "E:\AA\MiscVBAFunctions\test_data\GetAllFiles\folder1\empty file.txt" = AllFiles(2) = Pass
    Pass = "E:\AA\MiscVBAFunctions\test_data\GetAllFiles\folder1\folder1\empty file.xlsx" = AllFiles(3) = Pass
    Pass = "E:\AA\MiscVBAFunctions\test_data\GetAllFiles\folder2\empty file.docx" = AllFiles(4) = Pass
    Pass = "E:\AA\MiscVBAFunctions\test_data\GetAllFiles\folder2\folder1\folder1\empty file.txt" = AllFiles(5) = Pass

    Test_GetAllFilesRecursive = Pass
End Function
