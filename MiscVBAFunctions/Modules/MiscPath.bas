Attribute VB_Name = "MiscPath"
Option Explicit


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
    
    EvalPath = AbsolutePath(EvalPath, WB)
End Function


Function AbsolutePath(ByVal PathString As String, Optional WB As Workbook = Nothing) As String
    ' Convert the input Path string to an absolute path string.
    ' The result is normalised to contain only backslashes
    '
    ' Args:
    '   PathString: Path to be converted to an absolute path
    '   WB: The WorkBook that will be used to convert the PathString to an absolute Path
    '       if the PathString is a relative Path.
    '
    ' Returns:
    '   The Absolute Path as a string
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    PathString = ConvertToBackslashes(PathString)
    Dim IsNetwokDrive As Boolean
    IsNetwokDrive = False
    
    If IsAbsolutePath(PathString) Then
        If PathHasServer(PathString) Then
            ' fso.GetAbsolutePathName(...) doesn't work with network paths, so "x:\" gets added instead
            ' to ensure fso.GetAbsolutePathName(...) can function.
            ' The "x:\" gets replaced by "\\" in the last step
            PathString = "x:\" & Mid(PathString, 3)
            IsNetwokDrive = True
        
        ElseIf Left(PathString, 1) = "\" Then
            ' If PathString starts with a "\" and its not a network drive, Prepend drive letter only.
            ' fso.GetAbsolutePathName(...) points to ThisWorkbook's drive letter.
            PathString = PathGetDrive(WB.Path) & PathString
        End If
    Else
        ' Prepend WB.Path to PathString if not an absolute Path.
        PathString = Path(WB.Path, PathString)
    End If

    ' Breaking example: fso.GetAbsolutePathName("\\hello\world\\..\2")
    AbsolutePath = Fso.GetAbsolutePathName(PathString)
    If IsNetwokDrive Then
        ' Remove the "x:\" prefix and replace it with "\\" if the PathString is a network drive.
        AbsolutePath = "\\" & Mid(AbsolutePath, 4)
    End If

End Function


Public Function Path(ParamArray Args() As Variant) As String
    ' Combine paths or path segments.
    '
    ' It tries to follow the convention of Python's `pathlib.Path()` class. However,
    ' here we simply return a string, not an object, so we are forced to decide up
    ' front which path separators to use. We opt for using backslashes `\` only, to
    ' allow support for network paths like `\\server1\asdf` on Windows.
    '
    ' Args:
    '   Args:
    '     The folder paths and the name of folders or a file to be combined. May be
    '     provided as separate string arguments, or a single array or collection
    '     argument. If the first argument is an Array or Collection, the rest of the
    '     arguments will be ignored. Otherwise, each argument must be a string.
    '
    ' Returns:
    '   The combination of paths with valid path separators.
    '
    ' Examples:
    '   Using separate arguments (ParamArray):
    '     ? Path("a", "b")
    '     a/b
    '
    '   Using an array:
    '     ? Path(array("a", "b"))
    '     a/b
    '
    '   Using a collection:
    '     ? Path(col("a", "b"))
    '     a/b
    '
    '   See the unit tests for more examples.
    
    ' Process arguments
    Dim ArgsCollection As Collection
    Dim Arg As Variant
    Dim TmpArr() As Variant
    If TypeOf Args(0) Is Collection Then
        ' Path(col("a", "b"))
        Set ArgsCollection = Args(0)
    ElseIf IsArray(Args(0)) Then
        ' Path(array("a", "b"))
        Set ArgsCollection = Col()
        For Each Arg In Args(0)
            ArgsCollection.Add Arg
        Next Arg
    Else
        ' Path("a", "b")
        Set ArgsCollection = Col()
        For Each Arg In Args
            ArgsCollection.Add Arg
        Next Arg
    End If
    
    ' Check for empty string as input
    ' Return empty string witout go through the below
    If ArgsCollection.Count = 1 And ArgsCollection(1) = vbNullString Then Exit Function
    
    ' Always use backslash, because we need to support server paths like `\\server1\asdf` on Windows.
    ' This is unfortunate, because forward slashes `/` are the universal standard, and is even
    ' supported by Windows in most places, while backslash `\` works ONLY on Windows.
    Dim Slash As String
    Slash = "\"
    
    ' Collect path segments from args.
    Dim SegmentsRegex As Object
    Set SegmentsRegex = PathSegmentsRegex()
    Dim Segments As Collection
    Set Segments = Col()
    Dim I As Integer
    Dim SegmentMatches As Variant
    Dim SegmentMatch As Variant
    Dim LastKnownDrive As String
    LastKnownDrive = ""
    Dim ArgStr As String
    For I = 1 To ArgsCollection.Count
        ArgStr = ArgsCollection(I)
        
        If PathHasDrive(ArgStr) Then
            ' This is an absolute path with a drive letter.
            ' Throw away everything that came before it.
            Set Segments = Col()
            LastKnownDrive = PathGetDrive(ArgStr)
        ElseIf PathHasServer(ArgStr) Then
            ' This is a network path.
            ' Throw away everything that came before it, but preserve the extra leading
            ' backslash, which indicates a network path.
            Set Segments = Col("\")
        ElseIf PathStartsWithSlash(ArgStr) Then
            ' This is an absolute path without a drive letter.
            ' Throw away everything that came before it, but preserve the last known drive letter.
            Set Segments = Col(LastKnownDrive)
        Else
            ' This is a relative path. Continue collecting segments as normal.
        End If
        
        Set SegmentMatches = SegmentsRegex.Execute(ArgStr)
        For Each SegmentMatch In SegmentMatches
            Segments.Add SegmentMatch.Value
        Next SegmentMatch
    Next
    
    'Debug.Print "Segment: " & Segments(1)
    Path = Segments(1)
    For I = 2 To Segments.Count
        'Debug.Print "Segment: " & Segments(I)
        Path = Path & Slash & Segments(I)
    Next
End Function


Private Function ConvertToBackslashes(pth As String) As String
    ConvertToBackslashes = Replace(pth, "/", "\")

End Function


Public Function IsAbsolutePath(P As String) As Boolean
    ' Check if the given path is absolute. On Windows, this is when it starts with a forward slash, backslash, or drive letter.
    IsAbsolutePath = PathHasDrive(P) Or PathStartsWithSlash(P)
End Function


Public Function PathStartsWithSlash(P As String) As Boolean
    Dim Start As String
    Start = Left(P, 1)
    PathStartsWithSlash = (Start = "/") Or (Start = "\")
End Function


Public Function PathEndsWithSlash(P As String) As Boolean
    Dim Last As String
    Last = Right(P, 1)
    PathEndsWithSlash = (Last = "/") Or (Last = "\")
End Function


Public Function PathHasServer(P As String) As Boolean
    ' Check if the path has a server name at the start, e.g. `server1` in `\\server1\foo`.
    ' Note that Windows does not see `//server1/foo` as a network path, but we will allow it.
    PathHasServer = PathServerRegex().test(P)
End Function


Public Function PathGetServer(P As String) As String
    ' Get the server name at the start of a path, e.g. `server1` in `\\server1\foo`.
    ' Note that Windows does not see `//server1/foo` as a network path, but we will allow it.
    Dim Matches As Variant
    Set Matches = PathServerRegex().Execute(P)
    
    If Matches.Count <> 1 Then
        PathGetServer = ""
    Else
        PathGetServer = Matches.Item(0).Value
    End If
End Function


Public Function PathServerRegex() As Object
    ' Get a regular expression object to match the server name at the start of a path, e.g. `server1` in `\\server1\foo`.
    ' Note that Windows does not see `//server1/foo` as a network path, but we will allow it.
    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = "^[\/\\][\/\\]([^\/\\]+)"
    Regex.IgnoreCase = True
    Regex.Global = False
    Regex.MultiLine = False
    
    Set PathServerRegex = Regex
End Function


Public Function PathHasDrive(P As String) As Boolean
    PathHasDrive = PathDriveRegex().test(P)
End Function


Public Function PathGetDrive(P As String) As String
    Dim Matches As Variant
    Set Matches = PathDriveRegex().Execute(P)
    
    If Matches.Count <> 1 Then
        PathGetDrive = ""
    Else
        PathGetDrive = Matches.Item(0).Value
    End If
End Function


Public Function PathDriveRegex() As Object
    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = "^[a-z]:"
    Regex.IgnoreCase = True
    Regex.Global = False
    Regex.MultiLine = False
    
    Set PathDriveRegex = Regex
End Function


Public Function PathSegmentsRegex() As Object
    ' Get a regular expression that splits a path string into segments.
    Set PathSegmentsRegex = CreateObject("VBScript.RegExp")
    PathSegmentsRegex.Pattern = "[^\/\\]+"
    PathSegmentsRegex.IgnoreCase = True
    PathSegmentsRegex.Global = True
    PathSegmentsRegex.MultiLine = False
End Function
