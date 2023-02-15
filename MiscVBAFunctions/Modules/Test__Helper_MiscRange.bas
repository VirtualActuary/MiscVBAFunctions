Attribute VB_Name = "Test__Helper_MiscRange"
Option Explicit

Function Test_RangeToLO_fail()
    Const ExpectedError As Long = 58
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Dim RangeStart As Range
    Dim RangeTest As Range
    Dim Arr(1, 2) As Variant
    Dim LO As ListObject

    'Act:
    Set WB = ExcelBook("", False, False)
    Arr(0, 0) = "col1"
    Arr(0, 1) = "col2"
    Arr(0, 2) = "col3"
    Arr(1, 0) = "=[d]"
    Arr(1, 1) = "=d"
    Arr(1, 2) = 1
    
    
    Set RangeStart = WB.ActiveSheet.Range("B4")
    Set RangeTest = ArrayToRange(Arr, RangeStart, True)
    
    Set LO = RangeToLO(WB.ActiveSheet, RangeTest, "myTable")
    Set LO = RangeToLO(WB.ActiveSheet, RangeTest, "myTable")

    Test_RangeToLO_fail = False
    Exit Function

TestFail:
    If Err.Number = ExpectedError Then
        Test_RangeToLO_fail = True
    Else
        Test_RangeToLO_fail = False
    End If
End Function
