Attribute VB_Name = "MiscArray"
Option Explicit

Function ErrorToNullStringTransformation(tableArr As Variant) As Variant
    If is2D(tableArr) Then
        ErrorToNullStringTransformation = ErrorToNull2D(tableArr)
    Else
        ErrorToNullStringTransformation = ErrorToNull1D(tableArr)
    End If
End Function


Function EnsureDotSeparatorTransformation(tableArr As Variant) As Variant
    If is2D(tableArr) Then
        EnsureDotSeparatorTransformation = EnsureDotSeparator2D(tableArr)
    Else
        EnsureDotSeparatorTransformation = EnsureDotSeparator1D(tableArr)
    End If
End Function


Function DateToStringTransformation(tableArr As Variant) As Variant
    If is2D(tableArr) Then
        DateToStringTransformation = DateToString2D(tableArr)
    Else
        DateToStringTransformation = DateToString1D(tableArr)
    End If
End Function


' Check if a collection is 1D or 2D.
' 3D is not supported
Private Function is2D(arr As Variant)
    On Error GoTo Err
    is2D = (UBound(arr, 2) > 1)
    Exit Function
Err:
    is2D = False
End Function


Private Function dateToString(d As Date) As String
    If d = Int(d) Then ' no hours, etc:
        dateToString = Format(d, "yyyy-mm-dd")
    Else ' add hours and seconds - VBA can't keep more details in any case...
        dateToString = Format(d, "yyyy-mm-dd hh:mm:ss")
    End If
End Function


' Converts the decimal seperator in the float input to a "."
' and returns the result as a string.
' Only works when converting from the system's decimal seperator.
' Custom seperators not supported.
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
    Dim I As Long, J As Long
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
    Dim I As Long, J As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsNumeric(tableArr(I)) Then ' force numeric values to use . as decimal separator
            tableArr(I) = decStr(tableArr(I))
        End If
    Next I
    EnsureDotSeparator1D = tableArr
End Function


Private Function DateToString2D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsDate(tableArr(I, J)) Then ' format dates as strings to avoid some user's stupid default date settings
                tableArr(I, J) = dateToString(CDate(tableArr(I, J)))
            End If
        Next J
    Next I
    DateToStringTransformation = tableArr
End Function


Private Function DateToString1D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        If IsDate(tableArr(I)) Then ' format dates as strings to avoid some user's stupid default date settings
            tableArr(I) = dateToString(CDate(tableArr(I)))
        End If
    Next I
    DateToStringTransformation = tableArr
End Function
