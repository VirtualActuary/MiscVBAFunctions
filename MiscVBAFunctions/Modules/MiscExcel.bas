Attribute VB_Name = "MiscExcel"
Option Explicit

Public Function ExcelBook( _
      Path As String _
    , Optional MustExist As Boolean = False _
    , Optional ReadOnly As Boolean = False _
    , Optional SaveOnError As Boolean = False _
    , Optional CloseOnError As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    
    On Error GoTo finally
    
    If fso.FileExists(Path) Then
    
        Set ExcelBook = OpenWorkbook(Path, ReadOnly)
    
    Else
        
        If MustExist Then
            Err.Raise -999, , "FileNotFoundError: File '" & fso.GetAbsolutePathName(Path) & "' does not exist."
        Else
            Set ExcelBook = Workbooks.Add
            
            If SaveOnError Then
                ExcelBook.SaveAs Path
            End If
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
    
End Function

Function OpenWorkbook( _
      Path As String _
    , Optional ReadOnly As Boolean = False _
    ) As Workbook
    
    
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

