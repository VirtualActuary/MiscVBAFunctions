Attribute VB_Name = "MiscExcel"
Option Explicit

Private Sub ModuleInitialize()
    Dim WB As Workbook
    Set WB = ExcelBook(fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel23763464453.xlsx"), True)
    
End Sub

Public Function ExcelBook( _
      Path As String _
    , Optional MustExist As Boolean = False _
    , Optional ReadOnly As Boolean = False _
    , Optional SaveOnError As Boolean = False _
    , Optional CloseOnError As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    ' Create an Excel Workbook with custom arguments.
    '
    ' Args:
    '   Path: Path to the file.
    '   MustExist: If True, the file must exist. If it doesn't an error is raised.
    '   ReadOnly: If True, the file is opened in readOnly mode.
    '   SaveOnError: If True, the file is saved if an error is raised.
    '   CloseOnError: If True, close the file if an error was raised.
    '
    ' Returns:
    '   The created/opened Workbook.
    
    On Error GoTo finally
    If fso.FileExists(Path) Then
        Set ExcelBook = OpenWorkbook(Path, ReadOnly)
    Else
        Debug.Print "3", MustExist
        If MustExist Then
            'On Error GoTo 0
            Err.Raise -999, , "FileNotFoundError: File '" & fso.GetAbsolutePathName(Path) & "' does not exist."
        Else
            Set ExcelBook = Workbooks.Add
            
            'If SaveOnError Then
            ExcelBook.SaveAs Path
            'End If
        End If
        
    End If
    
    Exit Function
    
finally:
    If SaveOnError Then
        ExcelBook.Save
    End If
    
    If CloseOnError Then
        ExcelBook.Close (False)
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    
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



