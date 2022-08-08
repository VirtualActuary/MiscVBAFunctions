Attribute VB_Name = "MiscOs"
Option Explicit

Public Function Path(ParamArray Paths() As Variant) As String
    ' Combines folder paths and the name of folders or a file and
    ' returns the combination with valid path separators.
    ' Multiple Paths can be combined.
    '
    ' Args:
    '   entries: The folder paths and the name of folders or a file to be combined.
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


Public Function ExpandEnvironmentalVariables(pth As String) As String
    ' Find the Windows/Linux system environment variables in the input path
    ' and convert it to string in the input path and return it.
    '
    ' Args:
    '   pth: input Path
    '
    ' Returns:
    '   A string of Path with environment variables converted to its matching value.
    
    ExpandEnvironmentalVariables = pth
    If InStr(pth, "%") > 0 Then
        Dim I As Integer
        Dim EnvironVar As String
        Dim EnvironVarArr() As String
        EnvironVarArr = Split(pth, "%")
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


Private Function ConvertToBackslashes(pth As String) As String
    ConvertToBackslashes = Replace(pth, "/", "\")

End Function


Public Function EvalPath(pth As String, Optional WB As Workbook) As String
    ' Convert a path to absolute path.
    ' Converts system variables to String in the Path.
    ' Convert "/" to "\" in the Path.
    ' If the input Path doesn't start with "[A-Za-z]" and then ":" or starts with "\\",
    ' the selected WorkBook's Path is used to create the absolute path of the input Path.
    '
    ' Args:
    '   Pth: input path
    '   WB: Optional WorkBook.
    '
    ' Returns:
    '   The absolute Path.
    
    If WB Is Nothing Then Set WB = ThisWorkbook

    EvalPath = ExpandEnvironmentalVariables(pth)
    EvalPath = ConvertToBackslashes(EvalPath)

    If (Left(EvalPath, 1) Like "[A-Za-z]" And Mid(EvalPath, 2, 1) = ":") Or Left(EvalPath, 2) = "\\" Then
        EvalPath = fso.GetAbsolutePathName(EvalPath)
    Else
        EvalPath = fso.GetAbsolutePathName(Path(WB.Path, EvalPath))
    End If
End Function


Private Function parentDir(ByVal folder)
    parentDir = Left$(folder, InStrRev(folder, "\") - 1)
End Function


Public Sub CreateFolders(ByVal strPath As String, _
                  Optional doShell As Boolean = False)
    ' source: https://stackoverflow.com/questions/10803834/is-there-a-way-to-create-a-folder-and-sub-folders-in-excel-vba
    ' this code will make the folders recursively if required.
    ' the strPath can, therefore, be a Path where multiple sub-dirs don't exist yet.
    '
    ' Args:
    '   strPath: The desired directory
    '   doShell: if True, the directories will be created using WScript.shell

    Dim elm As Variant
    Dim strCheckPath As String
    Dim I As Integer
    Dim J As Integer
    Dim pathTempSplit() As String
    strPath = EvalPath(strPath)
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")

    Dim pathTemp As String ' for when the relative names becomes too long:
    pathTemp = EvalPath(strPath)

    If fso.FolderExists(pathTemp) Then Exit Sub
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
        If fso.FolderExists(strCheckPath) = False Then
            If doShell Then
                wsh.Run "cmd /c ""mkdir """ + strCheckPath + """"""
            Else
                On Error Resume Next
                MkDir strCheckPath
                On Error GoTo 0
            End If
        End If
    Next J
    If Not fso.FolderExists(pathTemp) Then
        Err.Raise 53, , "Could not create folder with path: " & pathTemp & ". Ensure you have write access to the required folder."
    End If
End Sub


Public Function RunShell(ByVal command As String, Optional WaitOnReturn As Boolean = True)
    ' run WScript.Shell with the selected command.
    '
    ' Args:
    '   command: Command to execute.
    '   WaitOnReturn: If True, The program will wait until this shell command is finishes before continuing.
    '
    ' Returns:
    '   The task ID of the started program.
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    'RunShell = wsh
    ':https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-function
    Dim windowStyle As Integer: windowStyle = 0 ' vbHide (no pop-ups)
    
    ' If the Shell function successfully executes the named file, it returns the task ID of the started program.
    RunShell = wsh.Run(command, windowStyle, WaitOnReturn)
End Function


