Attribute VB_Name = "MiscString"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function randomString(length As Variant)
    ' Create a random string containing hex characters only.
    ' (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F)
    '
    ' Args:
    '   length: Number of characters that the string must have.
    '
    ' Returns:
    '   The Random string.
    
    Dim S As String
    While Len(S) < length
        S = S & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(S, 1, length)
End Function


Public Function IsString(ByVal v As Variant) As Boolean
    ' Test is a Variant is of type String
    '
    ' Args:
    '   v: Variant variable to test
    '
    ' Returns
    '   Boolean indication if the input is a String
    IsString = False
    
    Dim S As String
    
    ' If not an object (s=v error test) then test if not numeric
    On Error Resume Next
        S = v
        If Err.Number = 0 Then
            If Not IsNumeric(v) Then IsString = True
        End If
End Function


Public Function EndsWith(StrComplete As String, Ending As String) As Boolean
    'Test if a string ends with an ending string
    '
    ' Args:
    '   StrComplete: Input string
    '   Ending: Section look for in the input string
    '
    ' Returns:
    '   True if the input string ends with the correct ending, False otherwise
    
     Dim endingLen As Integer
     endingLen = Len(Ending)
     EndsWith = (Right(Trim(UCase(StrComplete)), endingLen) = UCase(Ending))
End Function


Public Function StartsWith(Str As String, Start As String) As Boolean
    'Test if a string starts with an starting string
    '
    ' Args:
    '   StrComplete: Input string
    '   Ending: Section look for in the input string
    '
    ' Returns:
    '   True if the input string ends with the correct ending, False otherwise
    
     Dim startLen As Integer
     startLen = Len(Start)
     StartsWith = (Left(Trim(UCase(Str)), startLen) = UCase(Start))
End Function


Public Function DecStr(Num As Variant) As String
    ' Convert to string and ensure decimal point (for doubles)
    '
    ' Args:
    '   Num: Input value (likely numerical) to be converted to Str
    '
    ' Returns:
    '   String value converted from the input.
    
    DecStr = CStr(Num)

    'Frikin ridiculous loops for VBA
    If IsNumeric(Num) Then
       DecStr = Replace(DecStr, Format(0, "."), ".")
       ' Format(0, ".") gives the system decimal separator
    End If

End Function


Public Function strToNum(StrInput As String) As Double
    ' Convert string to numer and ensure decimal point.
    ' An error is raised when input is invalid.
    '
    ' Args:
    '   StrInput: Number input in string format
    '
    ' Returns:
    '   A Double value if the input is valid.
    
    StrInput = Replace(StrInput, ".", ",") 'why oh why VBA???
    On Error GoTo errorHandle
    strToNum = 0 + StrInput

    Exit Function
errorHandle:
    Err.Raise Err.Number, Err.Source, "Cannot convert """ & StrInput & """ to Double.", Err.HelpFile, Err.HelpContext

End Function


Public Function DeStringify(StrInput As String, Optional IgnoreNonStringified = False) As String
    ' Convert stringed string to string
    '
    ' Args:
    '   StrInput: String input that must be de-stringified
    '   IgnoreNonStringified: Don't raise an error if the input string doesn't start and end with quotes
    '
    ' Returns:
    '   The de-stringified string
    
    If StartsWith(StrInput, """") And EndsWith(StrInput, """") Then
        DeStringify = Mid(StrInput, 2, Len(StrInput) - 2)
        DeStringify = Replace(DeStringify, """""", """")

    ElseIf StartsWith(StrInput, "'") And EndsWith(StrInput, "'") Then
        DeStringify = Mid(StrInput, 2, Len(StrInput) - 2)
        DeStringify = Replace(DeStringify, "''", "'")

    Else
        If IgnoreNonStringified Then
            DeStringify = StrInput
        Else
            Err.Raise Err.valueError, , "Error in deStringify: x must start and end with with """
        End If
    End If

End Function



Public Function StrRepr(StrInput As Variant) As Variant
    ' Wrap into neat arg-like printable representation
    '
    ' Args:
    '   StrInput: String to get the printablt representation of.
    '
    ' Returns:
    '   Printable representation of the input string
    
    StrRepr = StrInput

    If VarType(StrRepr) = vbString Then
        StrRepr = Replace(StrRepr, """", """""")
        StrRepr = """" & StrRepr & """"
    End If

End Function

