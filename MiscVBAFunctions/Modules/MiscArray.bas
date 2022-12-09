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
    If Is1D(Arr) Then
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
    Dim Col1 As Collection
    Set Col1 = New Collection
    For Each CurrVal In Arr
        Col1.Add CurrVal
    Next
    Set ArrayToCollection = Col1
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
    
    If Is2D(tableArr) Then
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
    
    If Is2D(tableArr) Then
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

    If Is2D(tableArr) Then
        DateToStringTransformation = DateToString2D(tableArr, fmt)
    Else
        DateToStringTransformation = DateToString1D(tableArr, fmt)
    End If
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
    Dim foo As Variant
    foo = UBound(Arr, 2)
    Exit Function
Err:
    Is1D = True
End Function


Function ArrayGetDimension(Arr() As Variant) As Long
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
    ArrayGetDimension = I - 1
End Function


Public Function IsArrayAllocated(Arr() As Variant) As Boolean
    ' From: http://www.cpearson.com/excel/vbaarrays.htm
    ' this could contain some other useful array helpers...
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated
    ' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
    ' allocated.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.
    '
    ' This function is just the reverse of IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function only accepts arrays as inputs (variant previously).
    '
    ' Args:
    '   arr: array to check if allocated.
    '
    ' Returns:
    '   True if array is allocated, False otherwise.
    
    Dim N As Long
    On Error Resume Next

    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(Arr, 1)
    If (Err.Number = 0) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Under some circumstances, if an array
        ' is not allocated, Err.Number will be
        ' 0. To acccomodate this case, we test
        ' whether LBound <= Ubound. If this
        ' is True, the array is allocated. Otherwise,
        ' the array is not allocated.
        '''''''''''''''''''''''''''''''''''''''
        If LBound(Arr) <= UBound(Arr) Then
            ' no error. array has been allocated.
            IsArrayAllocated = True
        Else
            IsArrayAllocated = False
        End If
    Else
        ' error. unallocated array
        IsArrayAllocated = False
    End If

End Function


Function JaggedArrayToLO(Table() As Variant, LoName As String, WS As Worksheet, Optional StartR As Range) As ListObject
    ' uses a staggedArray (https://stackoverflow.com/questions/9435608/how-do-i-set-up-a-jagged-array-in-vba)
    ' created like this: https://stackoverflow.com/a/24584110/6822528
    ' assumes the first entry contains all the "columns"
    ' subsequent rows may have fewer columns, but not more
    '
    ' Args:
    '   Table: Array of Arrays input. Content to be inserted into the LO
    '   LoName: Desired List Object name
    '   WS: Selected WorkSheet.
    '   StartR: Optional - Start range in the WS.
    '
    ' Returns:
    '   A ListObject with the Array of arrays as its content
    
    If StartR Is Nothing Then Set StartR = ActiveCell
    Dim index As Integer, NrCols As Integer
    index = LBound(Table, 1)
    NrCols = UBound(Table(0), 1) - LBound(Table(0), 1) + 1
    
    Dim I As Integer, J As Integer
    For I = LBound(Table, 1) To UBound(Table, 1)
        For J = LBound(Table(I), 1) To UBound(Table(I), 1)
            If J > NrCols - 1 + index Then Err.Raise ErrNr.SubscriptOutOfRange, , "Subsequent rows of jagged Array may not have more columns than the header"
            
            StartR.Offset(I - index, J - index).Value = Table(I)(J)
        Next J
    Next I
    
    Dim TableR As Range
    Set TableR = StartR.Resize(UBound(Table, 1) - LBound(Table, 1) + 1, NrCols)
    
    Set JaggedArrayToLO = WS.ListObjects.Add(xlSrcRange, TableR, , xlYes)
    JaggedArrayToLO.Name = LoName
    
End Function


Private Function DateToString(D As Date, fmt As String) As String
    DateToString = Format(D, fmt)
End Function


Private Function DecStr(x As Variant) As String
     DecStr = CStr(x)

     'Frikin ridiculous loops for VBA
     If IsNumeric(x) Then
        DecStr = Replace(DecStr, Format(0, "."), ".")
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
                tableArr(I, J) = DecStr(tableArr(I, J))
            End If
        Next J
    Next I
    EnsureDotSeparator2D = tableArr
End Function


Private Function EnsureDotSeparator1D(tableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsNumeric(tableArr(I)) Then ' force numeric values to use . as decimal separator
            tableArr(I) = DecStr(tableArr(I))
        End If
    Next I
    EnsureDotSeparator1D = tableArr
End Function


Private Function DateToString2D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsDate(tableArr(I, J)) Then ' format dates as strings to avoid some user's stupid default date settings
                tableArr(I, J) = DateToString(CDate(tableArr(I, J)), fmt)
            End If
        Next J
    Next I
    DateToString2D = tableArr
End Function


Private Function DateToString1D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        If IsDate(tableArr(I)) Then ' format dates as strings to avoid some user's stupid default date settings
            tableArr(I) = DateToString(CDate(tableArr(I)), fmt)
        End If
    Next I
    DateToString1D = tableArr
End Function


Function IsInArray(Arr As Variant, ValueToBeFound) As Boolean
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
    
    Dim I As Integer
    For I = LBound(Arr) To UBound(Arr)
        If Arr(I) = ValueToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next I
    IsInArray = False
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


Function ArrayDuplicates(Arr(), Optional IncludePosition As Boolean = False) As Variant()
    ' Finds all duplicates in the input array and returns a new array with only duplicates.
    ' If no duplicates exist, an empty array is returned.
    ' If an entry has multiple duplicates in the input array, there will be multiple entries in the output array.
    '
    ' Args:
    '   Arr: Input array with potential duplicates.
    '   IncludePosition: Optional - If True, Include: "{Entry: IndexOfEntry}: " with the duplicate value.
    '        If False, only return the duplicate value.
    '
    ' Returns:
    '   The array with duplicates only, or an empty array if there are no duplicates.
    
    Dim Dups() As Variant, I As Long
    Dim D As Object
    Dim Duplicates As Collection
    Set D = CreateObject("Scripting.Dictionary")
    Set Duplicates = New Collection
    Dim PrependStr As String
    
    For I = LBound(Arr) To UBound(Arr)
        If D.Exists(Arr(I)) Then
            If IncludePosition Then
                Duplicates.Add "{Entry: " & I + 1 & "}: " & Arr(I)
            Else
                Duplicates.Add Arr(I)
            End If
        Else
            D.Add Item:=Arr(I), Key:=Arr(I)
        End If
        
    Next I
    
    If Duplicates.Count > 0 Then
        ReDim Dups(Duplicates.Count - 1)
        For I = 1 To Duplicates.Count
            Dups(I - 1) = Duplicates(I)
        Next I
    End If
    
    ArrayDuplicates = Dups
    
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
    
    Dimension = ArrayGetDimension(Arr)
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

