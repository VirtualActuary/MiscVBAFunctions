Attribute VB_Name = "MiscOs"
Option Explicit


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


Public Function ExpandEnvironmentalVariables(Pth As String) As String
    ' Find the Windows/Linux system environment variables in the input path
    ' and convert it to string in the input path and return it.
    '
    ' Args:
    '   pth: input Path
    '
    ' Returns:
    '   A string of Path with environment variables converted to its matching value.
    
    ExpandEnvironmentalVariables = Pth
    If InStr(Pth, "%") > 0 Then
        Dim I As Integer
        Dim EnvironVar As String
        Dim EnvironVarArr() As String
        EnvironVarArr = Split(Pth, "%")
        Dim IsEnvironmentalVariable As Boolean
        
        Do While UBound(EnvironVarArr) >= 2
            IsEnvironmentalVariable = False
            EnvironVar = Environ(EnvironVarArr(1))
            If Len(EnvironVar) Then
                IsEnvironmentalVariable = True
            End If

            If IsEnvironmentalVariable = True Then
                EnvironVarArr(0) = EnvironVarArr(0) & EnvironVar & EnvironVarArr(2)
                If UBound(EnvironVarArr) > 2 Then
                    For I = 1 To UBound(EnvironVarArr) - 2
                        EnvironVarArr(I) = EnvironVarArr(I + 2)
                    Next
                End If
                ReDim Preserve EnvironVarArr(UBound(EnvironVarArr) - 2)
                
            Else
                EnvironVarArr(0) = EnvironVarArr(0) & "%" & EnvironVarArr(1)
                For I = 1 To UBound(EnvironVarArr) - 1
                    EnvironVarArr(I) = EnvironVarArr(I + 1)
                Next
                ReDim Preserve EnvironVarArr(UBound(EnvironVarArr) - 1)
            End If
        Loop
        
        ' Add the last entries that can't be EnvironVar
        For I = 1 To UBound(EnvironVarArr)
            EnvironVarArr(0) = EnvironVarArr(0) & "%" & EnvironVarArr(I)
        Next
        ExpandEnvironmentalVariables = EnvironVarArr(0)
    End If
    
End Function


Public Function ParentDir(ByVal Folder)
    ParentDir = Left$(Folder, InStrRev(Folder, "\") - 1)
End Function


Public Sub CreateFolders( _
        ByVal StrPath As String, _
        Optional DoShell As Boolean = False _
    )
    If DoShell = True Then
        Err.Raise ErrNr.InternalError, , ErrorMessage(ErrNr.InternalError, "DoShell no longer supported.")
    End If
    
    Debug.Print "DeprecationWarning: CreateFolders is deprecated. Use 'MakeDirs' in the future instead."
    MakeDirs StrPath
                  
End Sub
                  
                  
Public Function MakeDirs(ByVal StrPath As String, Optional ExistOk As Boolean = True) As Folder
    ' source: https://stackoverflow.com/questions/10803834/is-there-a-way-to-create-a-folder-and-sub-folders-in-excel-vba
    ' this code will make the folders recursively if required.
    ' the StrPath can, therefore, be a Path where multiple sub-dirs don't exist yet.
    '
    ' Args:
    '   StrPath: The desired directory
   
    Dim StrCheckPath As String
    Dim I As Integer
    Dim J As Integer
    Dim PathTempSplit() As String
   
    Dim PathTemp As String ' for when the relative names becomes too long:
    PathTemp = EvalPath(StrPath)

    If Fso.FolderExists(PathTemp) Then
        ' if the folder already exists, no need to create anything
        If ExistOk = False Then
            Err.Raise ErrNr.FileAlreadyExists, , ErrorMessage(ErrNr.FileAlreadyExists, "Could not create folder with path: " & PathTemp & ". Folder already exists.")
        Else
            Set MakeDirs = Fso.GetFolder(PathTemp)
            Exit Function
        End If
    End If
    

    StrCheckPath = ""
    PathTempSplit = Split(PathTemp, "\")
    If Left(PathTemp, 2) = "\\" Then
        ' then it's an unmapped server
        ' we do not want to check if the server exists, so move to the next directory/folder
        For I = 0 To 2
            StrCheckPath = StrCheckPath & PathTempSplit(I) & "\"
        Next I
    End If

    For J = I To UBound(PathTempSplit)
        StrCheckPath = StrCheckPath & PathTempSplit(J) & "\"
        If Fso.FolderExists(StrCheckPath) = False Then
            On Error Resume Next
            MkDir StrCheckPath
            On Error GoTo 0
        End If
    Next J
    If Not Fso.FolderExists(PathTemp) Then
        Err.Raise ErrNr.FileNotFound, , ErrorMessage(ErrNr.FileNotFound, "Could not create folder with path: " & PathTemp & ". Ensure you have write access to the required folder.")

    End If
    Set MakeDirs = Fso.GetFolder(PathTemp)
End Function


Public Function RunShell(ByVal Command As String, Optional WaitOnReturn As Boolean = True)
    ' run WScript.Shell with the selected command.
    '
    ' Args:
    '   command: Command to execute.
    '   WaitOnReturn: If True, The program will wait until this shell command is finishes before continuing.
    '
    ' Returns:
    '   The task ID of the started program.
    
    Dim Wsh As Object
    Set Wsh = VBA.CreateObject("WScript.Shell")
    'RunShell = wsh
    ':https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-function
    Dim WindowStyle As Integer: WindowStyle = 0 ' vbHide (no pop-ups)
    
    ' If the Shell function successfully executes the named file, it returns the task ID of the started program.
    RunShell = Wsh.Run(Command, WindowStyle, WaitOnReturn)
End Function


