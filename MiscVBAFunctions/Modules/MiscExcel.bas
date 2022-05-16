Attribute VB_Name = "MiscExcel"
Option Explicit

Private Sub ModuleInitialize()
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel23763464453.xlsx"), True)
    
End Sub

Public Function ExcelBook( _
      Optional Path As String = "" _
    , Optional MustExist As Boolean = False _
    , Optional ReadOnly As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    ' Create an Excel Workbook with custom arguments.
    '
    ' Args:
    '   Path: Path to the file.
    '   MustExist: If True, the file must exist. If it doesn't an error is raised.
    '   ReadOnly: If True, the file is opened in readOnly mode.
    '
    ' Returns:
    '   The created/opened Workbook.
    
    If Len(Path) = 0 Then
        If MustExist Then
            Err.Raise -997, , "Temp file can't have MustExist = True."
        End If
        If ReadOnly Then
            Err.Raise -996, , "Temp file can't open in ReadOnly mode."
        End If
        
        Set ExcelBook = Workbooks.Add
        Exit Function
    End If
    
    If fso.FileExists(Path) Then
        Set ExcelBook = OpenWorkbook(Path, ReadOnly)
        Exit Function
    End If
    
    If MustExist Then
        Err.Raise -999, , "FileNotFoundError: File '" & fso.GetAbsolutePathName(Path) & "' does not exist."
    End If
    
    If ReadOnly Then
        Err.Raise -998, , "File must exist to open in ReadOnly mode: File '" & fso.GetAbsolutePathName(Path) & "' does not exist."
    End If
    
    Set ExcelBook = Workbooks.Add
    ExcelBook.SaveAs Path
    
End Function

Public Function OpenWorkbook( _
      Path As String _
    , Optional ReadOnly As Boolean = False _
    ) As Workbook
    ' Open a Workbook. An error is raised if a file with the same name is already open.
    ' If ReadOnly is True and the Workbook is already open but not in ReadOnly mode, an error is raised.
    '
    ' Args:
    '   Path: Path to the file that gets opened.
    '   ReadOnly: If True, the file gets opened in ReadOnly mode.
    '
    ' Returns:
    '   The opened Workbook.
    
    If hasKey(Workbooks, fso.GetFileName(Path)) Then
        Set OpenWorkbook = Workbooks(fso.GetFileName(Path))
        
        ' check if the workbook is actually the one specified in path
        ' use AbsolutePathName to remove any relative path references  (\..\ / \.\)
        If LCase(OpenWorkbook.FullName) <> LCase(fso.GetAbsolutePathName(Path)) Then
            Debug.Print fso.GetAbsolutePathName(Path)
            Err.Raise 457, , "Existing workbook with the same name is already open: '" & fso.GetFileName(Path) & "'"
        End If
        
        If ReadOnly And OpenWorkbook.ReadOnly = False Then
            Err.Raise -999, , "Workbook'" & fso.GetFileName(Path) & "' is already open and is not in ReadOnly mode. Only closed workbooks can be opened as readonly."
        End If
    Else
        Set OpenWorkbook = Workbooks.Open(Path, ReadOnly:=ReadOnly)
    End If
End Function



