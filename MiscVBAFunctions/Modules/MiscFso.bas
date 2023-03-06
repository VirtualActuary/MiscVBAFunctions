Attribute VB_Name = "MiscFSO"
Option Explicit
' allows us to use FSO functions anywhere in the project
' Use this link to see the available functions + documentation for the the fso object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'
' Additional FSO-related functions are added here, as well as wrapper functions of the FSO class
' where we want different/additional functionality.

Public Fso As New FileSystemObject


Public Function GetAllFilesRecursive(Directory As Folder) As Collection
    ' Get all files in the given directory and sub-directories and
    ' return a Collection with the File objects.
    '
    ' Args:
    '   Directory: The directory to get the files from.
    '
    ' Returns:
    '   A Collection with all the File objects.
    
    Set GetAllFilesRecursive = New Collection
    GetAllFilesHelper Directory, GetAllFilesRecursive
    
End Function


Private Sub GetAllFilesHelper(Directory As Folder, ListOfFiles As Collection)
    
    Dim F As File
    For Each F In Directory.Files
        ListOfFiles.Add F
    Next F
    
    Dim SubDir As Folder
    For Each SubDir In Directory.SubFolders
        GetAllFilesHelper SubDir, ListOfFiles
    Next SubDir
    
End Sub
