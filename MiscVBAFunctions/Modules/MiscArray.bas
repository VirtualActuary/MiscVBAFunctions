Attribute VB_Name = "MiscArray"
'@IgnoreModule ImplicitByRefModifier
Option Explicit


Public Function ArrayToRange(DataIncludingHeaders() As Variant, StartCell As Range, Optional EscapeFormulas As Boolean = False) As Range
    ' This function copies the input array's data to a Range object.
    ' Since the StartCell argument is connected to a WorkSheet, this data will be copied to the cells in that WorkSheet.
    ' If the input array isn't 2D, an error is raised.
    '
    ' Args:
    '   DataIncludingHeaders: Array containing the headers and data.
    '                         The first row is treated as the header.
    '   StartCell: Cell object. The Range will start at this cell.
    '              The WorkBook and WorkSheet connected to this object is used in this function.
    '   EscapeFormulas: Optional Boolean input.
    '                   If True, formulas get copied as text. (E.x. "=d" -> "'=d")
    '                   If False, the data is copied as is.
    '                   If this is False and "=[foo]" gets copied, the function will crash.
    '
    ' Returns:
    '   The Range of the data is returned, starting ath the StartCell cell.
    
    If Not is2D(DataIncludingHeaders) Then
        Err.Raise Number:=9, _
              Description:="2D array required."
    End If
    
    Dim StartRow As Long
    Dim EndRow As Long
    Dim StartColumn As Long
    Dim EndColumn As Long

    StartRow = StartCell.Row
    StartColumn = StartCell.Column
    EndRow = StartRow + UBound(DataIncludingHeaders) - LBound(DataIncludingHeaders)
    EndColumn = StartColumn + UBound(DataIncludingHeaders, 2) - LBound(DataIncludingHeaders, 2)

    Dim CellRange As Range
    Set CellRange = StartCell.Parent.Range(StartCell, StartCell.Parent.Cells(EndRow, EndColumn))


    Dim CountOuter As Long
    Dim CountInner As Long
    
    If EscapeFormulas Then
        For CountOuter = LBound(DataIncludingHeaders) To UBound(DataIncludingHeaders)
            For CountInner = LBound(DataIncludingHeaders, 2) To UBound(DataIncludingHeaders, 2)
                If Not IsError(DataIncludingHeaders(CountOuter, CountInner)) Then ' don't even try if it's an error value, else we get type mismatch
                    If Left(DataIncludingHeaders(CountOuter, CountInner), 1) = "=" Then
                        DataIncludingHeaders(CountOuter, CountInner) = "'" & DataIncludingHeaders(CountOuter, CountInner)
                    End If
                End If
            Next
        Next
    End If
    
    CellRange.Value = DataIncludingHeaders
    Set ArrayToRange = CellRange

End Function


Public Function ArrayToNewTable( _
        TableName As String, _
        DataIncludingHeaders() As Variant, _
        StartCell As Range, _
        Optional EscapeFormulas As Boolean = False _
    ) As ListObject
    ' This function creates a Table (as ListObject) and populates the table with the
    ' array's data (DataIncludingHeaders).
    ' If the input array isn't 2D, an error is raised.
    ' If the input table name (TableName) already exists in the WorkBook, an error is raised.
    '
    ' Args:
    '   TableName: Name of new table. Must be a unique name in the selected WB
    '   DataIncludingHeaders: Array containing the headers and data.
    '                         The first row is treated as the header.
    '   StartCell: Cell object. The Table will start at this cell.
    '              The WorkBook and WorkSheet connected to this object is used in this function.
    '   EscapeFormulas: Optional Boolean input.
    '                   If True, formulas get copied as text. (E.x. "=d" -> "'=d")
    '                   If False, the data is copied as is.
    '                   If this is False and "=[foo]" gets copied, the function will crash.
    '
    ' Returns:
    '   The Table is returned as a ListObject
    
    If HasLO(TableName, StartCell.Worksheet.Parent) Then
        Err.Raise Number:=-999, _
              Description:="Table name already exists."
    End If
    
    Dim CellRange As Range
    Set CellRange = ArrayToRange(DataIncludingHeaders, StartCell, EscapeFormulas)
    
    Set ArrayToNewTable = StartCell.Worksheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=CellRange, _
    xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium2")
    ArrayToNewTable.Name = TableName
End Function



Public Function ArrayToCollection(Arr() As Variant) As Collection
    ' Take an array as an input and return it as a Collection
    '
    ' Args:
    '   arr: Input array
    '
    ' Returns:
    '   Collection containing the values of the input array.
    
    Dim CurrVal As Variant
    Dim col1 As Collection
    Set col1 = New Collection
    For Each CurrVal In Arr
        col1.Add CurrVal
    Next
    Set ArrayToCollection = col1
End Function


Public Function ErrorToNullStringTransformation(tableArr() As Variant) As Variant
    ' Replaces all Errors in the input array with vbNullString.
    ' The input array is modified (pass by referance) and the function returns the array
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array that potentially contains error entries.
    '
    ' Returns:
    '   Array with the changed values.
    
    If is2D(tableArr) Then
        ErrorToNullStringTransformation = ErrorToNull2D(tableArr)
    Else
        ErrorToNullStringTransformation = ErrorToNull1D(tableArr)
    End If
End Function


Public Function EnsureDotSeparatorTransformation(tableArr() As Variant) As Variant
    ' Converts the decimal seperator in the float input to a "." for each entry in the input array
    ' and returns the result as an array of strings.
    ' Only works when converting from the system's decimal seperator.
    ' Custom seperators not supported.
    ' The input array is modified (pass by referance) and the function returns the array.
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array with float entries. Non numeric entries gets skipped.
    '
    ' Returns:
    '   Array with the changed string values.
    
    If is2D(tableArr) Then
        EnsureDotSeparatorTransformation = EnsureDotSeparator2D(tableArr)
    Else
        EnsureDotSeparatorTransformation = EnsureDotSeparator1D(tableArr)
    End If
End Function


Public Function DateToStringTransformation(tableArr() As Variant, Optional fmt As String = "yyyy-mm-dd") As Variant
    ' Converts all Date/DateTime entries in the input array to string.
    ' The input array is modified (pass by referance) and the function returns the array.
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array with potential Date/DateTime entries.
    '   fmt: String format of the date that it must convert to. Default = "yyyy-mm-dd"
    '
    ' Returns:
    '   Array where the Date/DateTime entries have been converted.

    If is2D(tableArr) Then
        DateToStringTransformation = DateToString2D(tableArr, fmt)
    Else
        DateToStringTransformation = DateToString1D(tableArr, fmt)
    End If
End Function


' Check if a collection is 1D or 2D.
' 3D is not supported
Private Function is2D(Arr As Variant)
    On Error GoTo Err
    is2D = (UBound(Arr, 2) > LBound(Arr, 2))
    Exit Function
Err:
    is2D = False
End Function


Private Function dateToString(d As Date, fmt As String) As String
    dateToString = Format(d, fmt)
End Function


Private Function decStr(x As Variant) As String
     decStr = CStr(x)

     'Frikin ridiculous loops for VBA
     If IsNumeric(x) Then
        decStr = Replace(decStr, Format(0, "."), ".")
        ' Format(0, ".") gives the system decimal separator
     End If

End Function


Private Function ErrorToNull2D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsError(tableArr(I, J)) Then ' set all error values to an empty string
                tableArr(I, J) = vbNullString
            End If
        Next J
    Next I
    ErrorToNull2D = tableArr
End Function


Private Function ErrorToNull1D(tableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsError(tableArr(I)) Then ' set all error values to an empty string
            tableArr(I) = vbNullString
        End If
    Next I
    ErrorToNull1D = tableArr
End Function


Private Function EnsureDotSeparator2D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsNumeric(tableArr(I, J)) Then ' force numeric values to use . as decimal separator
                tableArr(I, J) = decStr(tableArr(I, J))
            End If
        Next J
    Next I
    EnsureDotSeparator2D = tableArr
End Function


Private Function EnsureDotSeparator1D(tableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsNumeric(tableArr(I)) Then ' force numeric values to use . as decimal separator
            tableArr(I) = decStr(tableArr(I))
        End If
    Next I
    EnsureDotSeparator1D = tableArr
End Function


Private Function DateToString2D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsDate(tableArr(I, J)) Then ' format dates as strings to avoid some user's stupid default date settings
                tableArr(I, J) = dateToString(CDate(tableArr(I, J)), fmt)
            End If
        Next J
    Next I
    DateToString2D = tableArr
End Function


Private Function DateToString1D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        If IsDate(tableArr(I)) Then ' format dates as strings to avoid some user's stupid default date settings
            tableArr(I) = dateToString(CDate(tableArr(I)), fmt)
        End If
    Next I
    DateToString1D = tableArr
End Function
