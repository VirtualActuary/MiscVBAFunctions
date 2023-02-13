Attribute VB_Name = "MiscExcel"
Option Explicit

Private Sub ModuleInitialize()
    Dim WB As Workbook
    Set WB = ExcelBook(Fso.BuildPath(ThisWorkbook.Path, ".\tests\MiscExcel\MiscExcel23763464453.xlsx"), True)
    
End Sub

Public Function ExcelBook( _
      Optional Path As String = "" _
    , Optional MustExist As Boolean = False _
    , Optional ReadOnly As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    ' Create an Excel Workbook with custom arguments.
    ' If Path = "", a temp WB gets created.
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
    
    If Fso.FileExists(Path) Then
        Set ExcelBook = OpenWorkbook(Path, ReadOnly)
        Exit Function
    End If
    
    If MustExist Then
        Err.Raise -999, , "FileNotFoundError: File '" & Fso.GetAbsolutePathName(Path) & "' does not exist."
    End If
    
    If ReadOnly Then
        Err.Raise -998, , "File must exist to open in ReadOnly mode: File '" & Fso.GetAbsolutePathName(Path) & "' does not exist."
    End If
    
    Set ExcelBook = Workbooks.Add
    ExcelBook.SaveAs Path
    
End Function

Public Function OpenWorkbook( _
      Path As String _
    , Optional ReadOnly As Boolean = False _
    , Optional DisableUpdateLinksAndDisplayAlerts = True _
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
    
    If hasKey(Workbooks, Fso.GetFileName(Path)) Then
        Set OpenWorkbook = Workbooks(Fso.GetFileName(Path))
        
        ' check if the workbook is actually the one specified in path
        ' use AbsolutePathName to remove any relative path references  (\..\ / \.\)
        If VBA.LCase(OpenWorkbook.FullName) <> VBA.LCase(Fso.GetAbsolutePathName(Path)) Then
            Debug.Print Fso.GetAbsolutePathName(Path)
            Err.Raise 457, , "Existing workbook with the same name is already open: '" & Fso.GetFileName(Path) & "'"
        End If
        
        If ReadOnly And OpenWorkbook.ReadOnly = False Then
            Err.Raise -999, , "Workbook'" & Fso.GetFileName(Path) & "' is already open and is not in ReadOnly mode. Only closed workbooks can be opened as readonly."
        End If
    Else
        If DisableUpdateLinksAndDisplayAlerts Then
            Dim CurrentAskToUpdateLinks As Boolean
            Dim CurrentDisplayAlerts As Boolean
            CurrentAskToUpdateLinks = Application.AskToUpdateLinks
            CurrentDisplayAlerts = Application.DisplayAlerts
            Application.AskToUpdateLinks = False
            Application.DisplayAlerts = False
            
            Set OpenWorkbook = Workbooks.Open(Path, ReadOnly:=ReadOnly)
            
            Application.AskToUpdateLinks = CurrentAskToUpdateLinks
            Application.DisplayAlerts = CurrentDisplayAlerts
        Else
            Set OpenWorkbook = Workbooks.Open(Path, ReadOnly:=ReadOnly)
        End If
    End If
End Function


Public Function LastRow(WS As Worksheet) As Long
    ' Fetch the last row number that contains a value in any of the columns.
    ' Returns 0 if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last row number must be found
    '
    ' Returns:
    '   The last row number that contains a value in any of the columns.
    
    If WorksheetFunction.CountA(WS.Cells) = 0 Then
        LastRow = 0
        Exit Function
    End If
    LastRow = WS.Cells.Find(What:="*", After:=WS.Cells(1, 1), LookIn:=xlValues, lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function


Public Function LastColumn(WS As Worksheet) As Long
    ' Fetch the last column number that contains a value in any of the rows.
    ' Returns 0 if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column number must be found
    '
    ' Returns:
    '   The last column number that contains a value in any of the rows.
    
    If WorksheetFunction.CountA(WS.Cells) = 0 Then
        LastColumn = 0
        Exit Function
    End If
    LastColumn = WS.Cells.Find(What:="*", After:=WS.Cells(1, 1), LookIn:=xlValues, lookat:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

Public Function LastCell(WS As Worksheet) As Range
    ' Get the last active cell using LastRow() and LastColumn() and returns it as a range.
    ' This function doesn't fetch the last cell with containing a value, but rather returns the
    ' the cell with the last active row and last active column.
    ' An unset Range object is returned if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column must be found.
    '
    ' Returns:
    '   The last cell in a worksheet as a range.
    
    Dim Row As Long
    Dim Column As Long
    Row = LastRow(WS)
    Column = LastColumn(WS)
    If Row = 0 Or Column = 0 Then
        Exit Function
    End If
    
    Set LastCell = WS.Cells(Row, Column)
End Function


Public Function RelevantRange(WS As Worksheet) As Range
    ' Get the relevant range of a Worksheet. This relevant range always
    ' starts at Cell("A1") and ends at the cell with the last active row and last active column.
    ' An unset Range object is returned if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column must be found.
    '
    ' Returns:
    '   The range of active cells in the selected Worksheet.
    
    Dim LastCellEntry As Range
    Set LastCellEntry = LastCell(WS)
    
    If LastCellEntry Is Nothing Then
        Exit Function
    End If
    
    Set RelevantRange = WS.Range(WS.Cells(1, 1), LastCellEntry)
End Function


Public Function SanitiseExcelName(Name As String)
    ' Sanitises a proposed name to be a valid Excel name
    ' Any disallowed characters are replaced with `_`
    ' If the name starts with a number, `_` is prepended to the name
    ' Some documentation:
    ' https://support.microsoft.com/en-us/office/use-names-in-formulas-9cd0e25e-88b9-46e4-956c-cc395b74582a#:~:text=Guidelines%20for%20creating%20names
    '
    ' Args:
    '   Name: Proposed Excel name
    '
    ' Returns:
    '   A valid Excel name
    
    SanitiseExcelName = Name
    
    Dim Disallowed As String, I As Integer
    Disallowed = "- /*+=^!@#$%&?`~:;[](){}""'|,<>"
    For I = 1 To Len(Disallowed)
        If InStr(SanitiseExcelName, Mid(Disallowed, I, 1)) > 0 Then
            SanitiseExcelName = Replace(SanitiseExcelName, Mid(Disallowed, I, 1), "_")
        End If
    Next I
    
    If IsNumeric(Left(SanitiseExcelName, 1)) Then ' excel tables cannot start with a number
        SanitiseExcelName = "_" & SanitiseExcelName
    End If
    
End Function


Function VbaLocked() As Boolean
    ' Test if we have access to vba
    ' It's false if either protection <> pp_none or if user doesn't have access (an error)
    
    Dim VbProjectFlag As Boolean, VbaOpen As Boolean
    VbaOpen = True
    On Error GoTo VbProjectFlag_False
        If ThisWorkbook.VBProject.Protection <> vbext_ProjectProtection.vbext_pp_none Then
            VbaOpen = False
        End If
    GoTo VbProjectFlag_Keep

VbProjectFlag_False:
    VbaOpen = False
VbProjectFlag_Keep:

    VbaLocked = Not VbaOpen
End Function


Sub RenameSheet(SourceWS As Variant, NewSheetName As String, Optional RaiseErrorIfSheetNameExists = False)
    ' name a sheet given the proposed name (check first if it exists).
    ' Add "(NextAvailableNumber)" to the new sheet name if RaiseErrorIfSheetNameExists = False
    '
    ' Args:
    '   SourceWS: Worksheet whose name must be changed. This argument's Variable type can be String or Worksheet
    '   NewSheetName: Desired new worksheet name
    '   RaiseErrorIfSheetNameExists: Optional argument - If True, raise an error if the NewSheetName
    '       already exists in the WorkBook.
    
    If Not (VarType(SourceWS) = vbString Or VarType(SourceWS) = vbObject) Then
        Err.Raise ErrNr.TypeMismatch, , "Source Worksheet must be of type: Worksheet or String."
    End If
    
    Dim WS As Worksheet
    If VarType(SourceWS) = vbObject Then
        Set WS = SourceWS
    Else
        Set WS = GetLO(CStr(SourceWS))
    End If
        
    
    Dim SheetNames As New Collection
    Dim S As Worksheet
    For Each S In WS.Parent.Sheets
        SheetNames.Add S, S.Name
    Next S
    
    If RaiseErrorIfSheetNameExists Then
        If hasKey(SheetNames, NewSheetName) Then
            Err.Raise ErrNr.FileAlreadyExists, , "Worksheet name already exists."
        End If
    End If

    Dim Name As String
    Dim I As Integer
    I = 0
    Name = NewSheetName
    Do While hasKey(SheetNames, Name)
        I = I + 1
        Name = Left(NewSheetName, 25) & " (" & I & ")" ' 31 max characters - ie supports up to 999 sheets
    Loop
    WS.Name = Left(Name, 31)
End Sub


Function AddWS(ByVal Name As String, _
               Optional After As Worksheet, _
               Optional Before As Worksheet, _
               Optional WB As Workbook, _
               Optional RemoveGridLinesAndGreyFill As Boolean = True, _
               Optional ErrIfExists As Boolean = False, _
               Optional ForceNewIfExists As Boolean = False) As Worksheet
    ' Add a Worksheet to the selected Workbook.
    ' First check if the name already exists.
    '
    ' Args:
    '   Name: Name of the new Worksheet
    '   After: An object that specifies the sheet after which the new sheet is added.
    '   Before: An object that specifies the sheet before which the new sheet is added.
    '   WB: Workbook to add the Worksheet in.
    '   RemoveGridLinesAndGreyFill: If True, Remove gridlines and grey out the cells
    '   ErrIfExists: If True, Raise an error if the WorkSheet name already exists.
    '   ForceNewIfExists: If True, A unique key will be generated from the selected name
    '                     and will be used as the new name (e.x. MyTable -> MyTable1)
    '
    ' Returns:
    '   The new Worksheet object.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Name = Left(Name, 31)
    
    If hasKey(WB.Sheets, Name) Then
        If ErrIfExists Then
            Err.Raise ErrNr.FileAlreadyExists, , "Worksheet '" & Name & "' already exists"
        End If

        If ForceNewIfExists Then
            ' to allow new names up to *99 (99 sheets)
            Name = EnsureUniqueKey(WB.Sheets, Left(Name, 29))
        Else
            Set AddWS = WB.Sheets(Name)
            Exit Function
        End If
    End If
    
  
    If Not After Is Nothing Then
        Set AddWS = WB.Sheets.Add(After:=After)
    ElseIf Not Before Is Nothing Then
        Set AddWS = WB.Sheets.Add(Before:=Before)
    Else
        ' by default add to the last sheet
        Set AddWS = WB.Sheets.Add(After:=WB.Sheets(WB.Sheets.Count))
    End If
    
    AddWS.Name = Name
    
    If RemoveGridLinesAndGreyFill Then
        turnOffSheetGridLines AddWS
        makeDark -0.25, AddWS.Cells
    End If

End Function


Sub DeleteSheet(SheetName As String, Optional WB As Workbook = Nothing)
    ' Delete a sheet from the selected WorkBook.
    ' If the sheet doesn't exist, nothing happens.
    '
    ' Args:
    '   WB: Selected WorkBook. If left empty, ThisWorkbook is selected
    '   SheetName: Name of the sheet to search for.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim DA
    DA = Application.DisplayAlerts
    Application.DisplayAlerts = False
    If ContainsSheet(SheetName, WB) Then
        WB.Sheets(SheetName).Delete
    End If
    Application.DisplayAlerts = DA
End Sub


Function ContainsSheet(Key As Variant, Optional WB As Workbook = Nothing) As Boolean
    ' whether a sheet exists in a Workbook
    '
    ' Args:
    '   WB: Selected Workbook. If left empty, ThisWorkbook is selected
    '   Key: Sheet to search for.
    '
    ' Returns:
    '   True if the selected Workbook contains the sheet, False otherwise.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim obj As Variant
    Dim Sheets As Variant

    Set Sheets = WB.Sheets
    On Error GoTo Err
        ContainsSheet = True
        Set obj = Sheets(Key)
        Exit Function
Err:
        ContainsSheet = False
End Function


Sub InsertColumns(R As Range, Optional NrCols As Integer = 1)
    ' Insert 1 or more Columns to a Range.
    ' If the input Range object contains more than 1 cell, the first
    ' cell's location will be used to add the new column.
    '
    ' Args:
    '   R: Range object, used to place the new Column. This Range object will be altered
    '   NrCols: Number of columns to add.
    
    Dim I As Long
    For I = 1 To NrCols
        R.Cells(1, 1).EntireColumn.Insert
    Next I
End Sub


Sub InsertRows(R As Range, Optional NrRows As Integer = 1)
    ' Insert 1 or more rows to a Worksheet.
    ' If the input Range object contains more than 1 cell, the first
    ' cell's location will be used to add the new row.
    '
    ' Args:
    '   R: Range object, used to place the new row.
    '   NrCols: Number of rows to add.
    
    Dim I As Long
    For I = 1 To NrRows
        R.Cells(1, 1).EntireRow.Insert
    Next I
End Sub


