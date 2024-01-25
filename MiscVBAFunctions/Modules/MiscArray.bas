Attribute VB_Name = "MiscArray"
'@IgnoreModule ImplicitByRefModifier
Option Explicit


Public Function ArrayToRange( _
    Data() As Variant, _
    StartCell As Range, _
    Optional EscapeFormulas As Boolean = False, _
    Optional IncludesHeader As Boolean = False, _
    Optional PreventStringConversion As Boolean = True, _
    Optional NumberFormatPerColumn As Collection = Nothing _
) As Range
    ' This function copies data from the input array to a Range.
    '
    ' Data types are preserved.
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
    '     PreventStringConversion:
    '         If True, the number format of cells to which strings will be written are set to
    '         `@` before writing the data, to prevent the string from being converted to
    '         something else automatically. This could happen if the string looks like a date
    '         or a boolean, and is usually undesirable behavior.
    '         If False, the values are written to the cells without touching the number formats.
    '         This may be useful when the caller of this function already set the number
    '         formats of the destination range to something useful.
    '         Not compatible with the `NumberFormatPerColumn` argument.
    '     NumberFormatPerColumn:
    '         A number format for each column, in the same order as the columns in `Data`.
    '         Collection keys are ignored. If given, the data range of each column will be
    '         set to the corresponding number format from this collection before the values are
    '         written to the destination range.
    '
    ' Returns:
    '     The Range to which the data was written.
    
    If ArrayGetNumDimensions(Data) <> 2 Then
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage( _
            ErrNr.SubscriptOutOfRange, _
            "ArrayToRange can only function on 2D arrays. Use the `Ensure2dArray` function." _
        )
    End If
    
    Dim Sheet As Worksheet
    Set Sheet = StartCell.Parent
    
    Dim StartRow As Long
    StartRow = StartCell.Row
    
    Dim StartColumn As Long
    StartColumn = StartCell.Column
    
    Dim EndRow As Long
    EndRow = StartRow + UBound(Data) - LBound(Data)
    
    Dim EndColumn As Long
    EndColumn = StartColumn + UBound(Data, 2) - LBound(Data, 2)
    
    If IncludesHeader Then
        ' Format the header as `@`, because headers should always be strings.
        Dim HeaderRange As Range
        Set HeaderRange = Sheet.Range(StartCell, Sheet.Cells(StartRow, EndColumn))
        HeaderRange.NumberFormat = "@"
    End If
    
    Dim CellRange As Range
    Set CellRange = Sheet.Range(StartCell, Sheet.Cells(EndRow, EndColumn))
    
    Dim RowIndex As Long
    Dim ColumnIndex As Long
    
    If EscapeFormulas Then
        For RowIndex = LBound(Data) To UBound(Data)
            For ColumnIndex = LBound(Data, 2) To UBound(Data, 2)
                If Not IsError(Data(RowIndex, ColumnIndex)) Then ' don't even try if it's an error value, else we get type mismatch
                    If Left(Data(RowIndex, ColumnIndex), 1) = "=" Then
                        Data(RowIndex, ColumnIndex) = "'" & Data(RowIndex, ColumnIndex)
                        
                    End If
                    If IsNumeric(Data(RowIndex, ColumnIndex)) Then
                        If VarType(Data(RowIndex, ColumnIndex)) = VbString Then
                            Data(RowIndex, ColumnIndex) = "'" & Data(RowIndex, ColumnIndex)
                        End If
                    End If
                End If
            Next
        Next
    End If
    
    If PreventStringConversion Then
        For RowIndex = LBound(Data) To UBound(Data)
            For ColumnIndex = LBound(Data, 2) To UBound(Data, 2)
                If VarType(Data(RowIndex, ColumnIndex)) = VbString Then
                    StartCell.Offset(RowIndex, ColumnIndex).NumberFormat = "@"
                End If
            Next ColumnIndex
        Next RowIndex
    End If
    
    If EndRow > StartRow And Not NumberFormatPerColumn Is Nothing Then
        For ColumnIndex = LBound(Data, 2) To UBound(Data, 2)
            Sheet.Range(Sheet.Cells(StartRow + 1, StartColumn + ColumnIndex), Sheet.Cells(EndRow, StartColumn + ColumnIndex)).NumberFormat = NumberFormatPerColumn(ColumnIndex + 1)
        Next ColumnIndex
    End If
    
    CellRange.Value = Data
    Set ArrayToRange = CellRange

End Function


Public Function ArrayToNewTable( _
    TableName As String, _
    DataIncludingHeaders() As Variant, _
    StartCell As Range, _
    Optional EscapeFormulas As Boolean = False, _
    Optional NumberFormatPerColumn As Collection = Nothing _
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
    '     NumberFormatPerColumn:
    '         A number format for each column, in the same order as the columns in `Data`.
    '         Collection keys are ignored. If given, the data range of each column will be
    '         set to the corresponding number format from this collection before the values are
    '         written to the destination range.
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
        IncludesHeader:=True, _
        PreventStringConversion:=(NumberFormatPerColumn Is Nothing), _
        NumberFormatPerColumn:=NumberFormatPerColumn _
    )
    
    Set ArrayToNewTable = StartCell.Worksheet.ListObjects.Add( _
        SourceType:=XlSrcRange, _
        Source:=CellRange, _
        XlListObjectHasHeaders:=XlYes, _
        TablestyleName:="TableStyleMedium2" _
    )
    ArrayToNewTable.Name = TableName
End Function


Public Function Ensure2DArray(Arr() As Variant) As Variant()
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
    If ArrayGetNumDimensions(Arr) = 1 Then
        Dim I As Long
        ReDim ArrOut(0 To 0, 0 To UBound(Arr))
        For I = LBound(Arr) To UBound(Arr)
            ArrOut(0, I) = Arr(I)
        Next I
    Else
        ArrOut = Arr
    End If
    
    Ensure2DArray = ArrOut
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
    Dim Col1 As Collection
    Set Col1 = New Collection
    For Each CurrVal In Arr
        Col1.Add CurrVal
    Next
    Set ArrayToCollection = Col1
End Function


Public Function ErrorToNullStringTransformation(TableArr() As Variant) As Variant
    ' Replaces all Errors in the input array with vbNullString.
    ' The input array is modified (pass by referance) and the function returns the array
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array that potentially contains error entries.
    '
    ' Returns:
    '   Array with the changed values.
    
    Select Case ArrayGetNumDimensions(TableArr)
        Case 1
            ErrorToNullStringTransformation = ErrorToNull1D(TableArr)
        Case 2
            ErrorToNullStringTransformation = ErrorToNull2D(TableArr)
        Case Else
            Err.Raise ErrNr.SubscriptOutOfRange, , "Function only supports 1D and 2D arrays"
    End Select
    
End Function


Public Function EnsureDotSeparatorTransformation(TableArr() As Variant) As Variant
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
    
    Select Case ArrayGetNumDimensions(TableArr)
        Case 1
            EnsureDotSeparatorTransformation = EnsureDotSeparator1D(TableArr)
        Case 2
            EnsureDotSeparatorTransformation = EnsureDotSeparator2D(TableArr)
        Case Else
            Err.Raise ErrNr.SubscriptOutOfRange, , "Function only supports 1D and 2D arrays"
    End Select
    
End Function


Public Function DateToStringTransformation(TableArr() As Variant, Optional Fmt As String = "yyyy-mm-dd") As Variant
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

    Select Case ArrayGetNumDimensions(TableArr)
        Case 1
            DateToStringTransformation = DateToString1D(TableArr, Fmt)
        Case 2
            DateToStringTransformation = DateToString2D(TableArr, Fmt)
        Case Else
            Err.Raise ErrNr.SubscriptOutOfRange, , "Function only supports 1D and 2D arrays"
    End Select
End Function


Private Function Is2D(Arr As Variant)
    ' Check if a collection is 1D or 2D.
    ' 3D is not supported
    On Error GoTo Err
    Is2D = (UBound(Arr, 2) >= LBound(Arr, 2))
    Exit Function
Err:
    Is2D = False
End Function

Public Function Is1D(Arr As Variant)
    On Error GoTo Err
    Dim Foo As Variant
    Foo = UBound(Arr, 2)
    Exit Function
Err:
    Is1D = True
End Function


Function ArrayGetNumDimensions(Arr() As Variant) As Long
    ' Get the number of dimensions of an array.
    '
    ' Args:
    '   Arr: Array to get the dimensions of
    '
    ' Returns:
    '   Number of dimensions of the input array
    
    On Error GoTo Err
    Dim I As Long
    Dim Tmp As Long
    I = 0
    Do While True
        I = I + 1
        Tmp = UBound(Arr, I)
    Loop
Err:
    ArrayGetNumDimensions = I - 1
End Function


Private Function DateToString(D As Date, Fmt As String) As String
    DateToString = Format(D, Fmt)
End Function


Private Function ErrorToNull2D(TableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(TableArr, 1) To UBound(TableArr, 1)
        For J = LBound(TableArr, 2) To UBound(TableArr, 2)
            If IsError(TableArr(I, J)) Then ' set all error values to an empty string
                TableArr(I, J) = VbNullString
            End If
        Next J
    Next I
    ErrorToNull2D = TableArr
End Function


Private Function ErrorToNull1D(TableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(TableArr) To UBound(TableArr)
        If IsError(TableArr(I)) Then ' set all error values to an empty string
            TableArr(I) = VbNullString
        End If
    Next I
    ErrorToNull1D = TableArr
End Function


Private Function EnsureDotSeparator2D(TableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(TableArr, 1) To UBound(TableArr, 1)
        For J = LBound(TableArr, 2) To UBound(TableArr, 2)
            If IsNumeric(TableArr(I, J)) Then ' force numeric values to use . as decimal separator
                TableArr(I, J) = FixDecimalSeparator(TableArr(I, J))
            End If
        Next J
    Next I
    EnsureDotSeparator2D = TableArr
End Function


Private Function EnsureDotSeparator1D(TableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(TableArr) To UBound(TableArr)
        If IsNumeric(TableArr(I)) Then ' force numeric values to use . as decimal separator
            TableArr(I) = FixDecimalSeparator(TableArr(I))
        End If
    Next I
    EnsureDotSeparator1D = TableArr
End Function


Private Function DateToString2D(TableArr As Variant, Fmt As String) As Variant
    Dim I As Long, J As Long
    For I = LBound(TableArr, 1) To UBound(TableArr, 1)
        For J = LBound(TableArr, 2) To UBound(TableArr, 2)
            If IsDate(TableArr(I, J)) Then ' format dates as strings to avoid some user's stupid default date settings
                TableArr(I, J) = DateToString(CDate(TableArr(I, J)), Fmt)
            End If
        Next J
    Next I
    DateToString2D = TableArr
End Function


Private Function DateToString1D(TableArr As Variant, Fmt As String) As Variant
    Dim I As Long
    For I = LBound(TableArr, 1) To UBound(TableArr, 1)
        If IsDate(TableArr(I)) Then ' format dates as strings to avoid some user's stupid default date settings
            TableArr(I) = DateToString(CDate(TableArr(I)), Fmt)
        End If
    Next I
    DateToString1D = TableArr
End Function


Function IsInArray(Arr() As Variant, ValueToBeFound) As Boolean
    ' Source: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
    ' Check if a value is in the array.
    ' Not limited to string only.
    '
    ' Args:
    '   Arr: Input array
    '   ValueToBeFound: Value to look for in the array
    '
    ' Returns:
    '   True if value exists in the array, False otherwise.
    Dim Dimensions As Long
    Dimensions = ArrayGetNumDimensions(Arr)
    Dim I As Long
    
    If Dimensions = 1 Then
        
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) = ValueToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next I
        IsInArray = False
        
    ElseIf Dimensions = 2 Then
        Dim J As Long
        
        For I = LBound(Arr, 1) To UBound(Arr, 1)
            For J = LBound(Arr, 2) To UBound(Arr, 2)
                If Arr(I, J) = ValueToBeFound Then
                    IsInArray = True
                    Exit Function
                End If

            Next

        Next I
        IsInArray = False
    Else
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Only supported for 1D and 2D arrays.")
    End If
End Function


Function IsArrayUnique(Arr()) As Boolean
    ' Check if the array contains unique entries only
    '
    ' Args:
    '   Arr: Input array
    '
    ' Returns:
    '   True if all entries in the array is unique.
    
    IsArrayUnique = UBound(Arr) = UBound(ArrayUniqueValues(Arr))
End Function


Function ArrayUniqueValues(Arr() As Variant)  ' , Optional Dimension As Integer = 1
    ' Finds all unique entries in the input array and returns a new array with only these values.
    ' If no duplicates exist, a copy of the input array is returned.
    ' Works on 1D and 2D arrays
    '
    ' Args:
    '   Arr: Input array with potential duplicates.
    '   Dimension: 1 or 2. The dimension of the input array.
    '
    ' Returns:
    '   The array with unique entries only, or an empty array if there are no unique entries.
    
    Dim TmpArray() As Variant
    Dim M As Integer, N As Integer
    Dim Dimension As Long
    
    Dimension = ArrayGetNumDimensions(Arr)
    If Dimension = 2 Then
        ReDim TmpArray((UBound(Arr, 1) - LBound(Arr, 1) + 1) * (UBound(Arr, 2) - LBound(Arr, 2) + 1) - 1)
        
        For N = LBound(Arr, 1) To UBound(Arr, 1)
            For M = LBound(Arr, 2) To UBound(Arr, 2)
                TmpArray(N * (UBound(Arr, 2) - LBound(Arr, 2) + 1) + M) = Arr(N, M)
            Next M
        Next N
    Else
        ReDim TmpArray(UBound(Arr))
        For M = LBound(Arr) To UBound(Arr)
            TmpArray(M) = Arr(M)
        Next M
    End If
    
    Dim D As Object
    Set D = CreateObject("Scripting.Dictionary")
    
    Dim I As Long
    For I = LBound(TmpArray) To UBound(TmpArray)
        D(TmpArray(I)) = 1
    Next I
    
    Dim UniqueValues() As Variant
    ReDim UniqueValues(D.Count - 1)
    
    Dim J As Long
    J = 0
    Dim V As Variant
    For Each V In D.Keys()
        UniqueValues(J) = V
        J = J + 1
        'd.Keys() is a Variant array of the unique values in myArray.
        'v will iterate through each of them.
    Next V
    
    ArrayUniqueValues = UniqueValues

End Function
