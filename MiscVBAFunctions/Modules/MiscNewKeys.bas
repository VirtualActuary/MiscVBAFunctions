Attribute VB_Name = "MiscNewKeys"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function NewSheetName(Name As String, Optional WB As Workbook)
    ' this module is used to generate new keys to a container (collections, dict, sheets, etc)
    ' Use case is when we want to create a new sheet, but
    ' want to ensure we don't give a name that already exists in the workbook
    '
    ' Args:
    '   Name: Name of the Sheet.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   The unique name of the container.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' max 31 characters
    NewSheetName = Left(Name, 31)

    If Not HasKey(WB.Sheets, NewSheetName) Then
        ' sheet name doesn't exist, so we can continue
        Exit Function
    Else
        NewSheetName = GetNewKey(Name, WB.Sheets, 31)
    End If
End Function

Private Sub TestGetNewKey()

    Dim C As New Collection
    Dim I As Long
    
    C.Add "bla", "name"
    For I = 1 To 100
        C.Add "bla", "name" & I
    Next I
    
    Debug.Print GetNewKey("name", C), "name101"
    Debug.Print GetNewKey("NewName", C), "NewName"

End Sub


Public Function GetNewKey(Name As String, Container As Variant, Optional MaxLength As Long = -1, Optional Depth As Long = 0) As String
    ' get a key that does not exists in a container (dict or collection)
    ' we keep appending, 1, 2, 3, ..., 10, 11 until the key is unique
    ' MaxLength is used when the key has a restriction on the maximum length
    ' for example sheet names can only be 31 characters long
    '
    ' Args:
    '   Name: Name of the key
    '   Container: Container containing the existing keys
    '   MaxLength: Maximum length of the resulting key.
    '   depth: Starting number to append to the key, while searching for a unique key.
    '
    ' Returns:
    '   The unique key
    
    If MaxLength = -1 Then
        GetNewKey = Name
    Else
        GetNewKey = Left(Name, MaxLength)
    End If
    
    If Not HasKey(Container, GetNewKey) Then
        ' Key is "New" and we don't need further iteration
        Exit Function
    Else
        ' 31 max characters for sheet name
        Depth = Depth + 1
        If MaxLength = -1 Then
            GetNewKey = GetNewKey & Depth
        Else
            GetNewKey = Left(GetNewKey, MaxLength - Len(CStr(Depth))) & Depth
        End If
        
        If Not HasKey(Container, GetNewKey) Then
            Exit Function
        End If
        
        GetNewKey = GetNewKey(Name, Container, MaxLength, Depth)
    End If
End Function
