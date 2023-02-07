Attribute VB_Name = "MiscArray"
'@IgnoreModule ImplicitByRefModifier
Option Explicit


Public Function ArrayToRange( _
    Data() As Variant, _
    StartCell As Range, _
    Optional EscapeFormulas As Boolean = False, _
    Optional IncludesHeader As Boolean = False _
) As Range
    ' This function copies data from the input array to a Range.
    '
    ' Args:
    '     Data:
    '         Array containing the data. If this is not 2D, an error will be thrown.
    '         Use `Ensure2dArray` if the data might be 1D (e.g. a single row).
    '     StartCell:
    '         Cell object. The Range will start at this cell.
    '         The WorkBook and WorkSheet connected to this object is used in this function.
    '         Since this is tied to a WorkSheet, the data will be written to that sheet.
    '     EscapeFormulas:
    '         If True, formulas get copied as text. (E.x. "=d" -> "'=d")
    '         If False, the data is copied as is.
    '         If this is False and "=[foo]" gets copied, the function will crash.
    '     IncludesHeader:
    '         Whether `Data` includes a header. The header will be written with a number format
    '         of `@`, because headers must be strings, especially when creating `ListObject`s.
    '
    ' Returns:
    '     The Range to which the data was written.
    
    Dim StartRow As Long
    Dim EndRow As Long
    Dim StartColumn As Long
    Dim EndColumn As Long
    
    StartRow = StartCell.Row
    StartColumn = StartCell.Column
    EndRow = StartRow + UBound(Data) - LBound(Data)
    EndColumn = StartColumn + UBound(Data, 2) - LBound(Data, 2)
    
    If IncludesHeader Then
        ' Format the header as `@`, because headers should always be strings.
        Dim HeaderRange As Range
        Set HeaderRange = StartCell.Parent.Range(StartCell, StartCell.Parent.Cells(StartRow, EndColumn))
        HeaderRange.NumberFormat = "@"
    End If
    
    Dim CellRange As Range
    Set CellRange = StartCell.Parent.Range(StartCell, StartCell.Parent.Cells(EndRow, EndColumn))
    
    Dim CountOuter As Long
    Dim CountInner As Long
    
    If EscapeFormulas Then
        For CountOuter = LBound(Data) To UBound(Data)
            For CountInner = LBound(Data, 2) To UBound(Data, 2)
                If Not IsError(Data(CountOuter, CountInner)) Then ' don't even try if it's an error value, else we get type mismatch
                    If Left(Data(CountOuter, CountInner), 1) = "=" Then
                        Data(CountOuter, CountInner) = "'" & Data(CountOuter, CountInner)
                    End If
                End If
            Next
        Next
    End If
    
    CellRange.Value = Data
    Set ArrayToRange = CellRange

End Function


Public Function ArrayToNewTable( _
    TableName As String, _
    DataIncludingHeaders() As Variant, _
    StartCell As Range, _
    Optional EscapeFormulas As Boolean = False _
) As ListObject
    ' Create a ListObject and populate it with data from `DataIncludingHeaders`.
    '
    ' Args:
    '     TableName:
    '         Name of new table. Must be a unique name in the selected WB. If another table by
    '         this name already exists, an error is raised.
    '     DataIncludingHeaders:
    '         Array containing the headers and data. The first row is treated as the header,
    '         and will use `@` formatting. If this is not 2D, an error will be thrown.
    '         Use `Ensure2dArray` if the data might be 1D (e.g. a single row).
    '     StartCell:
    '         Cell object. The Range will start at this cell.
    '         The WorkBook and WorkSheet connected to this object is used in this function.
    '     EscapeFormulas:
    '         If True, formulas get copied as text. (E.x. "=d" -> "'=d")
    '         If False, the data is copied as is.
    '         If this is False and "=[foo]" gets copied, the function will crash.
    '
    ' Returns:
    '     The new ListObject.
    
    If HasLO(TableName, StartCell.Worksheet.Parent) Then
        Err.Raise Number:=-999, Description:="A table named '" & TableName & "' already exists."
    End If
    
    Dim CellRange As Range
    Set CellRange = ArrayToRange( _
        Data:=DataIncludingHeaders, _
        StartCell:=StartCell, _
        EscapeFormulas:=EscapeFormulas, _
        IncludesHeader:=True _
    )
    
    Set ArrayToNewTable = StartCell.Worksheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=CellRange, _
        xlListObjectHasHeaders:=xlYes, _
        tablestyleName:="TableStyleMedium2" _
    )
    ArrayToNewTable.Name = TableName
End Function


Public Function Ensure2dArray(Arr() As Variant) As Variant()
    ' Ensures an array is two-dimensional.
    '
    ' If the array is 1-dimensional, it will be used as the first row of the resulting 2-dimensional array.
    '
    ' Args:
    '     Arr: the input array.
    '
    ' Returns:
    '     The input array if it was already 2D, or a new 2D array if the original was 1D.
    
    Dim ArrOut() As Variant
    If is1D(Arr) Then
        Dim I As Long
        ReDim ArrOut(0 To 0, 0 To UBound(Arr))
        For I = LBound(Arr) To UBound(Arr)
            ArrOut(0, I) = Arr(I)
        Next I
    Else
        ArrOut = Arr
    End If
    
    Ensure2dArray = ArrOut
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


Private Function is2D(Arr As Variant)
    ' Check if a collection is 1D or 2D.
    ' 3D is not supported
    On Error GoTo Err
    is2D = (UBound(Arr, 2) >= LBound(Arr, 2))
    Exit Function
Err:
    is2D = False
End Function

Public Function is1D(Arr As Variant)
    On Error GoTo Err
    Dim foo As Variant
    foo = UBound(Arr, 2)
    Exit Function
Err:
    is1D = True
End Function

Private Function dateToString(d As Date, fmt As String) As String
    dateToString = Format(d, fmt)
End Function


Private Function decStr(X As Variant) As String
     decStr = CStr(X)

     'Frikin ridiculous loops for VBA
     If IsNumeric(X) Then
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
