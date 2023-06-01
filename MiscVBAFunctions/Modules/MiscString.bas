Attribute VB_Name = "MiscString"
'@IgnoreModule ImplicitByRefModifier
Option Explicit

Public Function RandomString(Length As Variant)
    ' Create a random string containing hex characters only.
    ' (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F)
    '
    ' Args:
    '   length: Number of characters that the string must have.
    '
    ' Returns:
    '   The Random string.
    
    Dim S As String
    While Len(S) < Length
        S = S & Hex(Rnd * 16777216)
    Wend
    RandomString = Mid(S, 1, Length)
End Function


Public Function IsString(ByVal V As Variant) As Boolean
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
        S = V
        If Err.Number = 0 Then
            If Not IsNumeric(V) Then IsString = True
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
    
     Dim EndingLen As Integer
     EndingLen = Len(Ending)
     EndsWith = (Right(Trim(UCase(StrComplete)), EndingLen) = UCase(Ending))
End Function


Public Function StartsWith(Str As String, Start As String) As Boolean
    ' Test if a string starts with an starting string
    '
    ' Args:
    '   Str: Input string
    '   Start: Section look for in the input string
    '
    ' Returns:
    '   True if the input string ends with the correct ending, False otherwise
    
     Dim StartLen As Integer
     StartLen = Len(Start)
     StartsWith = (Left(Trim(UCase(Str)), StartLen) = UCase(Start))
End Function


Public Function FixDecimalSeparator(X As Variant) As String
    ' Convert the input value to a string. If it's numeric, make sure that it uses "." as
    ' the decimal separator, regardless of the locale.
    '
    ' Args:
    '   X: The input value
    '
    ' Returns:
    '   A string representing a number, using "." as the decimal separator.
    
    FixDecimalSeparator = CStr(X)
    If IsNumeric(X) Then
        ' Format(0, ".") gives the system decimal separator
        FixDecimalSeparator = Replace(FixDecimalSeparator, Format(0, "."), ".")
    End If
End Function
