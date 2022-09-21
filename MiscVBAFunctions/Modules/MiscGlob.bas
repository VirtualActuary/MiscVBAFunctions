Attribute VB_Name = "MiscGlob"
Option Explicit


Public Function Glob(Dir As String, Pattern As String) As Collection
    ' Mimics Python's Glob function, but here the files are listed depth first,
    ' whereas in Python it's lists width first.
    '
    ' Available wildcards :  ?          :   Any single character
    '                        *          :   Zero or more characters
    '                        #          :   Any single digit (0–9)
    '                        [charlist] :   Any single character in charlist
    '                        [!charlist]:   Any single character not in charlist
    '
    ' Args:
    '   Dir: Base directory to find matching files from
    '   Pattern: Pattern to compare to the file names
    '
    ' Returns:
    '   A Collection with all files that match the input pattern
    
    If Len(Pattern) = 0 Then
        Err.Raise -345, , "ValueError: Unacceptable pattern: ''"
    End If
    
    Dim DirPath As folder
    Set DirPath = fso.GetFolder(Dir)
    
    Dim PatternSplitted() As String
    Pattern = Replace(Pattern, "/", "\")
    PatternSplitted = Split(Pattern, "\")

    Dim AllPaths As Collection
    If InStr(Pattern, "**") Then
        Set AllPaths = GetAllPaths(DirPath)
    Else
        Set AllPaths = GetAllPaths(DirPath, UBound(PatternSplitted))
    End If
    
    Dim NumPaths As Long
    NumPaths = AllPaths.Count
    
    Dim CurrentPath As String
    
    Dim I As Long
    I = 1
    While I <= NumPaths
        CurrentPath = CreateRelativePath(CStr(DirPath), CStr(AllPaths(I)))
        If IsGlobValid(CurrentPath, PatternSplitted, Dir) = False Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If

        I = I + 1
    Wend
    
    Set Glob = AllPaths
    
End Function



Public Function RGlob(Dir As String, Pattern As String) As Collection
    ' Mimics Python's RGlob function, but here the files are listed depth first,
    ' whereas in Python it's lists width first.
    '
    ' Available wildcards :  ?          :   Any single character
    '                        *          :   Zero or more characters
    '                        #          :   Any single digit (0–9)
    '                        [charlist] :   Any single character in charlist
    '                        [!charlist]:   Any single character not in charlist
    '
    ' Args:
    '   Dir: Base directory to find matching files from
    '   Pattern: Pattern to compare to the file names
    '
    ' Returns:
    '   A Collection with all files that match the input pattern
    
    
    Dim DirPath As folder
    Set DirPath = fso.GetFolder(Dir)
    
    If Pattern = "" Then
        Pattern = "**"
    End If
    
    Dim PatternSplitted() As String
    PatternSplitted = Split(Pattern, "\")

    Dim AllPaths As Collection
  
    Set AllPaths = GetAllPaths(DirPath)
    
    Dim NumPaths As Long
    NumPaths = AllPaths.Count
    
    Dim CurrentPath As String
    
    Dim I As Long
    I = 1
    While I <= NumPaths
        CurrentPath = CreateRelativePath(CStr(DirPath), CStr(AllPaths(I)))
  
        If IsRGlobValid(CurrentPath, PatternSplitted, Dir) = False Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If

        I = I + 1
    Wend
    
    Set RGlob = AllPaths
    
End Function


Private Function CreateRelativePath(DirPath As String, CurrentPath As String) As String
    Dim J As Long
    Dim CurrentPathSections() As String
    CurrentPathSections = Split(CurrentPath, "\")
    
    For J = UBound(Split(DirPath, "\")) + 1 To UBound(Split(CurrentPath, "\"))
        If Len(CreateRelativePath) Then
            CreateRelativePath = CreateRelativePath + "\" + CurrentPathSections(J)
        Else
            CreateRelativePath = CurrentPathSections(J)
        End If
    Next
End Function


Private Function IsGlobValid(CurrentPath As String, PatternSplitted() As String, BaseDir As String)
    
    Dim PathSplitted() As String
    PathSplitted = Split(CurrentPath, "\")
    
    Dim PathLen As Long
    PathLen = UBound(PathSplitted)
    
    Dim PatternLen As Long
    PatternLen = UBound(PatternSplitted)

    Dim RelativePattern As String
    Dim RelativePath As String

    Dim I As Long

    For I = 0 To PatternLen
        
        If PatternSplitted(I) = "**" Then
            
            If I = PatternLen Then
                If fso.FolderExists(BaseDir + "\" + CurrentPath) Then
                    IsGlobValid = True
                Else
                    IsGlobValid = False
                End If
                Exit Function
            End If

            Dim ArrRest() As String
            ReDim ArrRest(PatternLen - I - 1) As String
            Dim J As Long
            
            For J = I + 1 To PatternLen
                ArrRest(J - I - 1) = PatternSplitted(J)
            Next J

            If RelativePath <> "" Then
                IsGlobValid = IsRGlobValid(Mid(CurrentPath, Len(RelativePath) + 1), ArrRest, BaseDir + "\" + RelativePath)
            Else
                IsGlobValid = IsRGlobValid(CurrentPath, ArrRest, BaseDir)
            End If
            Exit Function
        End If
        
        If I > PathLen Then
            IsGlobValid = False
            Exit Function
        End If
        
        If Len(RelativePath) Then
            RelativePath = RelativePath + "\" + PathSplitted(I)
            RelativePattern = RelativePattern + "\" + PatternSplitted(I)
        Else
            RelativePath = PathSplitted(I)
            RelativePattern = PatternSplitted(I)
        End If
        
        If Not RelativePath Like RelativePattern Then
            IsGlobValid = False
            Exit Function
        
        Else
            If I = PatternLen Then
                IsGlobValid = True
                Exit Function
            End If
        End If
    Next I
End Function


Private Function IsRGlobValid(ByVal CurrentPath As String, PatternSplit() As String, BaseDir As String)
    ' must still test for scenerios where there are multiple recursions (not at beginning or end)
    
    Dim PatternSplitted() As String
    PatternSplitted = PatternSplit
    
    If CurrentPath = "" Then
        If (UBound(PatternSplitted) = 0 And PatternSplitted(0) = "**") Then
            IsRGlobValid = True
        Else
            IsRGlobValid = False
        End If
        Exit Function
    End If
    
    Dim PathSplitted() As String
    PathSplitted = Split(CurrentPath, "\")
    PathSplitted = ReverseArray(PathSplitted)
    PatternSplitted = ReverseArray(PatternSplitted)
    
    Dim PathLen As Long
    PathLen = UBound(PathSplitted)
    
    Dim PatternLen As Long
    PatternLen = UBound(PatternSplitted)

    Dim RelativePattern As String
    Dim RelativePath As String

    Dim I As Long

    For I = 0 To PatternLen
        If I > PathLen Then
            IsRGlobValid = False
            Exit Function
        End If
    
        If PatternSplitted(I) = "**" Then
            Dim ArrRest() As String
            
            If I = 0 Then
                ' if The last section of the pattern is "**", only folders are considered.
                If Not fso.FolderExists(BaseDir + "\" + CurrentPath) Then
                    IsRGlobValid = False
                    Exit Function
                End If
            End If
            If I = PatternLen Then
                ' All other checks passed and this is final check, so test passed.
                IsRGlobValid = True
                Exit Function
            End If
                
            ReDim ArrRest(PatternLen - I - 1) As String
            Dim J As Long
            
            For J = I + 1 To PatternLen
                ArrRest(J - I - 1) = PatternSplitted(J)
            Next J
            
            IsRGlobValid = IsRecursiveRGlobValid(Left(CurrentPath, Len(CurrentPath) - Len(RelativePath)), ArrRest)
            Exit Function
        End If
        
        If Len(RelativePath) Then
            RelativePath = RelativePath + "\" + PathSplitted(I)
            RelativePattern = RelativePattern + "\" + PatternSplitted(I)
        Else
            RelativePath = PathSplitted(I)
            RelativePattern = PatternSplitted(I)
        End If
        
        If Not RelativePath Like RelativePattern Then
            IsRGlobValid = False
            Exit Function
        Else
            If I = PatternLen Then
                IsRGlobValid = True
                Exit Function
            End If
        End If
    Next I
    
End Function


Private Function IsRecursiveRGlobValid(CurrentPath As String, PatternSplitted() As String) As Boolean
    ' Glob for multiple recursions
    ' Example: pattern = "foo\**\bar\**\tree
    
    Dim PathSplitted() As String
    PathSplitted = Split(CurrentPath, "\")
    
    Dim RelativePattern As String
    Dim RelativePath As String
    
    Dim PatternLen As Long
    PatternLen = UBound(PatternSplitted)

    Dim I As Long
    Dim J As Long
    
    ' must look at all different Path sections.
    For I = 0 To UBound(PathSplitted)
        RelativePattern = ""
        For J = 0 To UBound(PatternSplitted)
            ' If pattern is longer that remaining path
            If I + UBound(PatternSplitted) > UBound(PathSplitted) Then
                IsRecursiveRGlobValid = False
                Exit Function
            End If
            
            ' Build a relative path to be compared to Pattern section
            If RelativePath = "" Then
                RelativePath = PathSplitted(I + J)
            Else
                RelativePath = RelativePath + "\" + PathSplitted(I + J)
            End If
            
            ' If current pattern section is "**" call this function again (recursion)
            If PatternSplitted(J) = "**" Then
                ' Untested section
                If J = UBound(PatternSplitted) Then
                    IsRecursiveRGlobValid = True
                    Exit Function
                End If
                
                Dim ArrRest() As String
                ReDim ArrRest(PatternLen - J - 1) As String
                Dim K As Long
                
                For K = I + 1 To PatternLen
                    ArrRest(K - J - 1) = PatternSplitted(K)
                Next K
            
                If IsRecursiveRGlobValid(Mid(CurrentPath, Len(RelativePath) + 1), ArrRest) Then
                    IsRecursiveRGlobValid = True
                    Exit Function
                End If
                GoTo NextIteration
            End If
            
            ' Build relative pattern
            If RelativePattern = "" Then
                RelativePattern = PatternSplitted(J)
            Else
                RelativePattern = RelativePattern + "\" + PatternSplitted(J)
            End If
            If J = UBound(PatternSplitted) Then
                If RelativePath Like "*" & RelativePattern & "*" Then
                    IsRecursiveRGlobValid = True
                    Exit Function
                End If
            Else
                If Not RelativePath Like RelativePattern Then
                    GoTo NextIteration
                End If
                 
            End If
            
        Next J
        
NextIteration:
    Next I
    
    IsRecursiveRGlobValid = False
End Function


Private Function ReverseArray(arr() As String) As String()
    Dim StartPos As Long
    Dim EndPos As Long
    StartPos = LBound(arr)
    EndPos = UBound(arr)
    Dim ReverseArr() As String
    ReDim ReverseArr(StartPos To EndPos)

    Dim I As Long
    Dim Counter As Long
    Counter = StartPos
    For I = EndPos To StartPos Step -1
        ReverseArr(Counter) = arr(I)
        Counter = Counter + 1
    Next I
    ReverseArray = ReverseArr
End Function


Private Function GetAllPaths(Directory As folder, Optional MaxDepth = 999, Optional CurrentDepth = 0) As Collection

    Set GetAllPaths = New Collection
    GetAllPaths.Add Directory
    GetAllPathsHelper Directory, GetAllPaths, MaxDepth

End Function


Private Sub GetAllPathsHelper(Directory As folder, ListOfFiles As Collection, Optional MaxDepth = 999, Optional CurrentDepth = 0)
    ' Depth first ordering

    Dim F As File
    For Each F In Directory.Files
        ListOfFiles.Add F
    Next F

    Dim SubDir As folder
    For Each SubDir In Directory.SubFolders
        ListOfFiles.Add SubDir
        If CurrentDepth < MaxDepth Then
            GetAllPathsHelper SubDir, ListOfFiles, MaxDepth, CurrentDepth + 1
        End If
    Next SubDir

End Sub

