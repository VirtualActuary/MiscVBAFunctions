Attribute VB_Name = "MiscDate"
Option Explicit



Function EoWeek(StartDate As Variant, Weeks As Double) As Variant
    ' Like EOMONT but for weeks
    ' Gives the date of the end of the week (Sunday) from the input date
    ' and offset of x weeks.
    ' Args:
    '   StartDate: Input date
    '   Weeks: Number of offset weeks
    '
    ' Returns:
    '   Date of the End of the week.

    With Application.WorksheetFunction
        EoWeek = .EDate(StartDate, 0) + 7 + (7 * Weeks) - .Weekday(StartDate, 2)
    End With
End Function
