Attribute VB_Name = "MiscOs"
Option Explicit

Public Function Path(ParamArray Paths() As Variant) As String
    ' Combines folder pathes and the name of folders or a file and
    ' returns the combination with valid path separators.
    '
    ' Args:
    '   entries: The folder pathes and the name of folders or a file to be combined.
    '
    ' Returns:
    '   The combination of paths with valid path separators.
    
    Dim Entry As Variant
    
    Path = Paths(0)

    For Entry = LBound(Paths) + 1 To UBound(Paths)
        Path = fso.BuildPath(Path, Paths(Entry))
    Next
    
End Function


Public Function Is64BitXl() As Boolean
    ' Check if the current version of Excel is 64-bit.
    '
    ' Returns:
    '   True if Excel is 64-bit. False if not.
    
    #If Win64 Then
        Is64BitXl = True
    #Else
        Is64BitXl = False
    #End If
End Function


Public Function OfficeBitness() As String
    ' Check if Office is 32- or 64-bit.
    
    ' Returns:
    '   "64-bit" if Office is 64-bit. Else "32-bit"
    
    OfficeBitness = IIf(Is64BitXl, "64-bit", "32-bit")
End Function


Public Function OpenXLFile() As Workbook
    ' open Excel file from dialog box (if InputsWB needed in playground routines):
    ' https://stackoverflow.com/a/39485571/6822528
    '
    ' Returns:
    '   The opened Excel file.
    
    Dim WB            As Workbook
    Dim FileName      As String
    Dim fd            As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Title = "Select the inputs file"

    ' Optional properties: Add filters
    fd.Filters.Clear
    fd.Filters.Add "Excel files", "*.xls*" ' show Excel file extensions only

    ' means success opening the FileDialog
    If fd.Show = -1 Then
        FileName = fd.SelectedItems(1)
    End If

    ' error handling if the user didn't select any file
    If FileName = "" Then
        Err.Raise Err.Number, Err.Source, "No Excel file was selected !", Err.HelpFile, Err.HelpContext
        End
    End If

    Set OpenXLFile = Workbooks.Open(FileName)

End Function


Public Function GetFilesInFolder(FolderPath As String, Optional FullPath As Boolean = False) As Collection
    ' Gets all the files in the selected folder and return the Names/Paths in a Collection.
    '
    ' Args:
    '   folderPath: Path to the folder.
    '   fullPath: True to collect the full paths. False to collect the names of the files only.
    '
    ' Returns:
    '   A Collection of the file names/paths.
    
    Dim folder As folder, file As file

    Set GetFilesInFolder = New Collection
    Set folder = fso.GetFolder(FolderPath)
    For Each file In folder.Files
        If FullPath Then
            GetFilesInFolder.Add file.Path
        Else
            GetFilesInFolder.Add file.Name
        End If
    Next file

End Function


Public Function GetFoldersInFolder(FolderPath As String, Optional FullPath As Boolean = False) As Collection
    ' Gets all the folders in the selected folder and return the Names/Paths in a Collection.
    '
    ' Args:
    '   folderPath: Path to the folder.
    '   fullPath: True to collect the full paths. False to collect the names of the folders only.
    '
    ' Returns:
    '   A Collection of the folder names/paths.
    
    Dim folder As folder, folder_i As folder

    Set GetFoldersInFolder = New Collection
    Set folder = fso.GetFolder(FolderPath)
    For Each folder In folder.SubFolders
        If FullPath Then
            GetFoldersInFolder.Add folder.Path
        Else
            GetFoldersInFolder.Add folder.Name
        End If
    Next folder

End Function


Public Function EvalPath(pth As String, Optional WB As Workbook) As String
    ' Convert relative to absolute path.
    '
    ' Args:
    '   pth: input path
    '   WB: Optional WorkBook.
    '
    ' Returns:
    '   The absolute Path.
    
    If WB Is Nothing Then Set WB = ThisWorkbook

    EvalPath = pth

    If InStr(EvalPath, "%") > 0 Then
        Dim environVar As String
        environVar = Split(EvalPath, "%")(1)

        On Error GoTo notenviron: ' if environ(environvar doesn't work its not an environVar)
            EvalPath = Replace(EvalPath, "%" & environVar & "%", Environ(environVar), Compare:=vbTextCompare)
notenviron:
    End If

    ' this also allows for ..\..\
    If Left(EvalPath, 2) = ".." Then
        EvalPath = ParentDir(WB.Path) & right(EvalPath, Len(EvalPath) - 2)
    ElseIf Left(EvalPath, 1) = "." Then
        EvalPath = WB.Path & right(EvalPath, Len(EvalPath) - 1)
    End If

End Function


Public Function ParentDir(ByVal folder)
    ' The the parent directory of the selected directory.
    '
    ' Args:
    '   folder: The input directory.
    '
    ' Returns:
    '   The parent directory
    
    ParentDir = Left$(folder, InStrRev(folder, "\") - 1)
End Function


Public Function IsInGitRepo(Optional WB As Workbook) As Boolean
    ' Test if the WorkBook is in a Git repo.
    '
    ' Args:
    '   WB: Optional Workbook.
    '
    ' Returns:
    '   True if the WorkBook is in a git repo. False otherwise.
    
    If WB Is Nothing Then Set WB = ThisWorkbook

    Dim Path As String, path_arr() As String, I As Integer
    path_arr = Split(WB.Path, "\")

    For I = UBound(path_arr) To LBound(path_arr) Step -1
        ReDim Preserve path_arr(I)
        Path = Join(path_arr, "\") & "\.git"
        If folderExists(Path) Then
            IsInGitRepo = True
            Exit Function
        End If
    Next I

End Function


Public Function GetGitRepoRootFolder(Optional WB As Workbook) As String
    ' Get the Root directory of the Git repo.
    '
    ' Args:
    '   WB: Optional WorkBook
    '
    ' Returns:
    '   String of the Root directory. Nothing is returned if there is no Git repo.
    
    If WB Is Nothing Then Set WB = ThisWorkbook

    Dim Path As String, path_arr() As String, I As Integer
    path_arr = Split(WB.Path, "\")

    For I = UBound(path_arr) To LBound(path_arr) Step -1
        ReDim Preserve path_arr(I)
        Path = Join(path_arr, "\") & "\.git"
        If folderExists(Path) Then
            GetGitRepoRootFolder = Join(path_arr, "\")
            Exit Function
        End If
    Next I

End Function


Public Function MakePath(Dir As String, _
                  FileName As String) As String
    ' Make a path from the input Dir and fileName.
    ' irrespective of whether directory contains the closing "\" or fileName starts with "\".
    ' Filename can be any path/ file to append to Dir
    '
    ' Args:
    '   Dir: The Directory
    '   fileName: File name. This can be a file/folder/path.
    '
    ' Returns:
    '   The string of the resulting Path.

    MakePath = IIf(right(Dir, 1) = "\", Left(Dir, Len(Dir) - 1), Dir) & "\" & IIf(Left(FileName, 1) = "\", Mid(FileName, 2), FileName)
End Function


Public Function RelativePath(fromPath As String, _
                      toPath As String, _
                      Optional fromIsFile As Boolean = True) As String
    ' Get the relative path from the 2 input directories.
    ' some of the logic from here:
    ' https://stackoverflow.com/a/3054692/6822528
    '
    ' Args:
    '   fromPath: The Path from which to get the relative path.
    '   toPath: The destination Path for the relative path.
    '
    ' Returns:
    '   The relative path.
    
    Dim common() As String, commonIndex As Integer
    Dim from() As String, to_() As String

    from = Split(fromPath, "\")
    to_ = Split(toPath, "\")

    commonIndex = 0
    Do While UBound(from) > commonIndex And UBound(to_) > commonIndex And _
             from(commonIndex) = to_(commonIndex)
        ReDim Preserve common(commonIndex)
        common(commonIndex) = from(commonIndex)
        commonIndex = commonIndex + 1
    Loop

    If commonIndex = 0 Then
        ' no common path
        RelativePath = toPath
        Exit Function
    End If

    Dim nrDirsUp As Integer

    If fromIsFile Then
        nrDirsUp = UBound(from) - commonIndex
    Else
        nrDirsUp = UBound(from) - commonIndex - 1
    End If

    Dim I As Integer
    If nrDirsUp = 0 Then
        RelativePath = "\."
    Else
        For I = 1 To nrDirsUp
            RelativePath = RelativePath & "\.."
        Next I
    End If
    RelativePath = Mid(RelativePath, 2)

    For I = commonIndex To UBound(to_)
        RelativePath = RelativePath & "\" & to_(I)
    Next I

End Function


Public Sub DeleteFileIfExists(Path)
    ' Delete the selected file if it exists.
    '
    ' Args:
    '   path: the input path.
    
    If file_Exists(Path) Then
        Kill Path
    End If
End Sub


Public Sub CreateFolders(ByVal strPath As String, _
                  Optional doShell As Boolean = False)
    ' source: https://stackoverflow.com/questions/10803834/is-there-a-way-to-create-a-folder-and-sub-folders-in-excel-vba
    ' this code will make the folders recursively if required
    '
    '
    ' Args:
    '   strPath: The desired directory
    '   doShell: if True, the directories will be created using WScript.shell

    Dim elm As Variant
    Dim strCheckPath As String
    Dim I As Integer
    Dim J As Integer
    Dim pathTempSplit() As String
    strPath = localFullName(strPath)

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")

    Dim pathTemp As String ' for when the relative names becomes too long:
    pathTemp = fso.GetAbsolutePathName(strPath)

    If folderExists(pathTemp) Then Exit Sub
        ' if the folder already exists, no need to create anything

    strCheckPath = ""

    pathTempSplit = Split(pathTemp, "\")
    If Left(pathTemp, 2) = "\\" Then
        ' then it's an unmapped server
        ' we do not want to check if the server exists, so move to the next directory/folder
        For I = 0 To 2
            strCheckPath = strCheckPath & pathTempSplit(I) & "\"
        Next I
    End If

    For J = I To UBound(pathTempSplit)
        strCheckPath = strCheckPath & pathTempSplit(J) & "\"
        If folderExists(strCheckPath) = False Then
            If doShell Then
                wsh.Run "cmd /c ""mkdir """ + strCheckPath + """"""
            Else
                On Error Resume Next
                MkDir strCheckPath
                On Error GoTo 0
            End If
        End If
    Next J
    If Not folderExists(pathTemp) Then
        Err.Raise userError.FileNotFound, , ErrorMessage(userError.FileNotFound, "Could not create folder with path: " & pathTemp & ". Ensure you have write access to the required folder.")
    End If
End Sub


Public Sub MakeFolderHidden(Path As String)
    ' make a file or folder hidden
    '
    ' Args:
    '   path: Path to directory to make hidden.
    
    Dim F As Object
    Set F = fso.GetFolder(Path)
    F.Attributes = 2

End Sub


Public Sub CopyAndRenameFile(sourceFilePath As String, targetFilePath As String)
    ' use the absolute path name so that the name is shorter
    ' in some instances where C:\foo\foo\..\..\foo\foo.xlsb is longer than
    ' 260 characters, this functions gives an error because the "filename" is too long
    '
    ' Args:
    '   sourceFilePath: source path
    '   targetFilePath: destination path

    Dim targetFileTemp As String
    targetFileTemp = fso.GetAbsolutePathName(targetFilePath)

    If Len(targetFileTemp) > 260 Then
        Err.Raise userError.BadFileNameOrNumber, , _
            ErrorMessage(userError.BadFileNameOrNumber, "File " & targetFileTemp & " is too long. Choose a shorter output_folder.")
    End If

    fso.CopyFile sourceFilePath, targetFileTemp

End Sub


Public Sub MoveAndRenameFile(sourceFilePath As String, targetFilePath As String)
    ' first check if target file exists, and if so delete it
    '
    ' Args:
    '   sourceFilePath: Source file
    '   targetFilePath: target file

    If FileExists(targetFilePath) = True Then Kill targetFilePath
    fso.MoveFile sourceFilePath, targetFilePath
End Sub


Public Sub CopyFolderAndSubfolders(ByVal Source As String, ByVal Destination As String, Optional overwrite As Boolean = True)
    ' Copy folder and subfolders. ByVal to not change the original text strings
    '
    ' Args:
    '   source: Directory to copy
    '   destination: Destination directory
    '   overwrite: True to overwrite existing files. False otherwise.

    If right(Source, 1) = "\" Then Source = Left(Source, Len(Source) - 1)
    If right(Destination, 1) = "\" Then Destination = Left(Destination, Len(Destination) - 1)

    Call fso.CopyFolder(Source, Destination, overwrite)

End Sub


Public Sub DeleteFolders(ByVal folderspec As String, Optional force As Boolean = False)
    ' ByVal to not change the original text strings
    ' deletes the folder selected
    '
    ' Args:
    '   folderspec: folder to delete.
    '   force: Force delete

    If right(folderspec, 1) = "\" Then folderspec = Left(folderspec, Len(folderspec) - 1)

    Call fso.DeleteFolder(folderspec, force)

End Sub


Private Sub test_deleteFolderContents()
    CreateFolders (EvalPath(".\Temp\tmp\TMP"))
    DeleteFolderContents EvalPath(".\Temp\")
End Sub


Public Sub DeleteFolderContents(ByVal folderspec As String)
    ' Delete the content in the selected folder.
    '
    ' Args:
    '   folderspec: Folder to delete from.

    Dim F As Object

    For Each F In fso.GetFolder(folderspec).Files
        F.Delete force:=True
    Next F

    For Each F In fso.GetFolder(folderspec).SubFolders
        F.Delete force:=True
    Next F

End Sub


Public Function GetDecimalSeparatorSettings() As Collection
    ' get the current settings in a keyed collection
    '
    ' Returns:
    '   A collection of the decimal seperator settings.
    
    Set GetDecimalSeparatorSettings = New Collection
    GetDecimalSeparatorSettings.Add Application.UseSystemSeparators, key:="UseSystemSeparators"
    GetDecimalSeparatorSettings.Add Application.DecimalSeparator, key:="DecimalSeparator"
End Function


Public Sub SetDecimalSeparatorSettings(settings As Collection)
    ' set Decimal Separator Settings
    '
    ' Args:
    '   settings: the selected Decimal Separator Settings
    
    Application.UseSystemSeparators = settings("UseSystemSeparators")
    Application.DecimalSeparator = settings("DecimalSeparator")
End Sub


Public Sub UseDotDecimalSep()
    ' use dot decimal separator
    ' then functions also use ',' to separate arguments instead of ';'

    Application.UseSystemSeparators = False
    Application.DecimalSeparator = "."
End Sub


Public Function RunShell(ByVal command As String)
    '   run WScript.Shell with the selected command.
    '
    ' Args:
    '   command: Command to execute.
    '
    ' Returns:
    '   The task ID of the started program.
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    RunShell = wsh.Run
    ':https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-function
    Dim windowStyle As Integer: windowStyle = 0 ' vbHide (no pop-ups)
    Dim waitOnReturn As Boolean: waitOnReturn = True ' wait for script to finish

    ' If the Shell function successfully executes the named file, it returns the task ID of the started program.
    RunShell = wsh.Run(command, windowStyle, waitOnReturn)
    Set wsh = Nothing
End Function


Public Sub RunCmd(ByVal command As String)
    ' Run a cmd command from runShell.
    '
    ' Args:
    '   command. Command to execute.
    
    RunShell "cmd /c """ & command & """"
End Sub


Public Function UserName()
    '   Get the Windows userName from environ$
    '
    ' Returns:
    '   The Windows UserName
    
    UserName = Environ$("UserName")
End Function

