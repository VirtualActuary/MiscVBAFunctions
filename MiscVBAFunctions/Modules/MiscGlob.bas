Attribute VB_Name = "MiscGlob"
Option Explicit


Function testrglob()
    Dim RemainingPath As Variant
    Dim c As Collection
    
    Set c = rglob("C:\AA\test_dir", "2021\**")
    
    For Each RemainingPath In c
        Debug.Print RemainingPath
    Next
End Function


Function testglob()
    Dim RemainingPath As Variant
    Dim c As Collection
    
    Set c = glob("C:\AA\test_dir", "**\2021\*")
    
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

    Set AllPaths = GetAllPaths(DirPath, LenPatternSplits)
    NumPaths = AllPaths.Count
    
    I = 1
    While I <= NumPaths
        If Not IsPathValid(CStr(AllPaths(I)), Pattern, UBound(Split(Dir, "\"))) Then
            AllPaths.Remove I
            NumPaths = NumPaths - 1
            I = I - 1
        End If
        I = I + 1
    Wend
    
    Set FindCorrectPaths = AllPaths

End Function


Public Function glob(Dir As String, Pattern As String) As Collection
    ' Limitations:
    '   Max 999 sub-dirs deep
    
    If Len(Pattern) = 0 Then
        Err.Raise -345, , "ValueError: Unacceptable pattern: ''"
    End If
    
    Dim LenPatternSplits As Long
    LenPatternSplits = UBound(Split(Pattern, "\"))
    If InStr(Pattern, "**") Then
        Pattern = Left(Pattern, InStr(Pattern, "**") - 1) + Mid(Pattern, InStr(Pattern, "**") + 3)
        LenPatternSplits = 999
    End If
    
    Set glob = FindCorrectPaths(Dir, Pattern, LenPatternSplits)
End Function


Public Function rglob(Dir As String, Pattern As String) As Collection
    ' "2021\**" this pattern doesn't work to run recursively again
    
    Set rglob = FindCorrectPaths(Dir, Pattern, 999)
End Function


'Function IsPathValid(CurrentPath As String, Pattern As String, DirPathLen As Long) As Boolean
'    IsPathValid = False
'    Dim LenPatternSplits As Long
'    LenPatternSplits = UBound(Split(Pattern, "\"))
'    Dim PathSplitted() As String
'    PathSplitted = Split(CurrentPath, "\")
'
'    Dim RelevantPathSection As String
'
'    If UBound(PathSplitted) - LenPatternSplits <= DirPathLen Then
'        ' Pattern can't go further back than the dir Path
'        Exit Function
'    End If
'
'    Dim I As Long
'    For I = DirPathLen + 1 To DirPathLen + 1 + LenPatternSplits
'        If Len(RelevantPathSection) Then
'            RelevantPathSection = RelevantPathSection & "\" & PathSplitted(I)
'        Else
'            RelevantPathSection = PathSplitted(I)
'        End If
'
'    Next
'
'    If RelevantPathSection Like Pattern Then
'        IsPathValid = True
'    End If
'
'End Function


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
    For I = UBound(PathSplitted) - LenPatternSplits To UBound(PathSplitted)
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



