Attribute VB_Name = "MiscGlob"
Option Explicit


Function testrglob()
    Dim RemainingPath As Variant
    Dim c As Collection
    
    Set c = rglob("C:\AA\test_dir", "NB\*.xlsx")
    
    For Each RemainingPath In c
        Debug.Print RemainingPath
    Next
End Function


Function testglob()
    Dim RemainingPath As Variant
    Dim c As Collection
    
    Set c = glob("C:\AA\test_dir", "")
    
    For Each RemainingPath In c
        Debug.Print RemainingPath
    Next
End Function
Function wut()
    Debug.Print InStr("sdfg", "fg")
End Function


Function FindCorrectPaths(Dir As String, Pattern As String, LenPatternSplits As Long) As Collection
    Dim AllPaths As Collection
    Dim DirPath As folder
    Dim NumPaths As Long
    Dim I As Long
    Set DirPath = fso.GetFolder(Dir)
    
'    Dim LenPatternSplits As Long
'    LenPatternSplits = UBound(Split(Pattern, "\"))
    
    Set AllPaths = GetAllPaths(DirPath, LenPatternSplits)
    NumPaths = AllPaths.Count
    
    I = 1
    While I <= NumPaths
        If Not IsPathValidRecursive(CStr(AllPaths(I)), Pattern, UBound(Split(Dir, "\"))) Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If
        I = I + 1
    Wend
    
    Set glob = AllPaths

End Function


Public Function glob(Dir As String, Pattern As String) As Collection
    ' Limitations:
    '   ** can't be called from glob. Use rglob instead.
    '   Max 999 sub-dirs deep
    
    If Not Len(Pattern) Then
        Err.Raise -345, , "ValueError: Unacceptable pattern: ''"
    End If
    
    Dim AllPaths As Collection
    Dim DirPath As folder
    Dim NumPaths As Long
    Dim I As Long
    Set DirPath = fso.GetFolder(Dir)
    
    Dim LenPatternSplits As Long
    LenPatternSplits = UBound(Split(Pattern, "\"))
    If InStr(Pattern, "**") Then
        LenPatternSplits = 999
    End If
    
    Set AllPaths = GetAllPaths(DirPath, LenPatternSplits)
    NumPaths = AllPaths.Count
    
    I = 1
    While I <= NumPaths
        If Not IsPathValidRecursive(CStr(AllPaths(I)), Pattern, UBound(Split(Dir, "\"))) Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If
        I = I + 1
    Wend
    
    Set glob = AllPaths

End Function


Public Function rglob(Dir As String, Pattern As String) As Collection

    Dim AllPaths As Collection
    Dim DirPath As folder
    Dim NumPaths As Long
    Dim I As Long
    Set DirPath = fso.GetFolder(Dir)
    Set AllPaths = GetAllPaths(DirPath)
    NumPaths = AllPaths.Count
    I = 1
    
    While I <= NumPaths
        If Not IsPathValidRecursive(CStr(AllPaths(I)), Pattern, UBound(Split(Dir, "\"))) Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If
        I = I + 1
    Wend
    
    Set rglob = AllPaths

End Function


Function IsPathValid(CurrentPath As String, Pattern As String, DirPathLen As Long) As Boolean
    IsPathValid = False
    Dim LenPatternSplits As Long
    LenPatternSplits = UBound(Split(Pattern, "\"))
    Dim PathSplitted() As String
    PathSplitted = Split(CurrentPath, "\")
    
    Dim RelevantPathSection As String
    
    If UBound(PathSplitted) - LenPatternSplits <= DirPathLen Then
        ' Pattern can't go further back than the dir Path
        Exit Function
    End If
    
    Dim I As Long
    For I = DirPathLen + 1 To DirPathLen + 1 + LenPatternSplits
        If Len(RelevantPathSection) Then
            RelevantPathSection = RelevantPathSection & "\" & PathSplitted(I)
        Else
            RelevantPathSection = PathSplitted(I)
        End If
        
    Next

    If RelevantPathSection Like Pattern Then
        IsPathValid = True
    End If
    
End Function


Function IsPathValidRecursive(CurrentPath As String, Pattern As String, DirPathLen As Long) As Boolean
    IsPathValidRecursive = False
    Dim LenPatternSplits As Long
    LenPatternSplits = UBound(Split(Pattern, "\"))
    Dim PathSplitted() As String
    PathSplitted = Split(CurrentPath, "\")
    
    Dim RelevantPathSection As String
    
    If UBound(PathSplitted) - LenPatternSplits <= DirPathLen Then
        ' Pattern can't go further back than the dir Path
        Exit Function
    End If
    
    Dim I As Long
    For I = UBound(PathSplitted) - LenPatternSplits To UBound(PathSplitted)
        If Len(RelevantPathSection) Then
            RelevantPathSection = RelevantPathSection & "\" & PathSplitted(I)
        Else
            RelevantPathSection = PathSplitted(I)
        End If
        
    Next

    If RelevantPathSection Like Pattern Then
        IsPathValidRecursive = True
    End If
    
End Function


Private Function GetAllPaths(Directory As folder, Optional MaxDepth = 999, Optional CurrentDepth = 0) As Collection
    Set GetAllPaths = New Collection
    GetAllPathsHelper Directory, GetAllPaths, MaxDepth
    
End Function


Private Sub GetAllPathsHelper(Directory As folder, ListOfFiles As Collection, Optional MaxDepth = 999, Optional CurrentDepth = 0)
    
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



