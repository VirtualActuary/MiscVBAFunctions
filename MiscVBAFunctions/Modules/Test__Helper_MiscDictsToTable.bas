Attribute VB_Name = "Test__Helper_MiscDictsToTable"
Option Explicit

Function Test_DictsToTable_fail_1(TableDict As Collection, RangeObj As Range)
    Const ExpectedError As Long = -997
    On Error GoTo TestFail
    
    DictsToTable TableDict, RangeObj, "someName"
    
    Test_DictsToTable_fail_1 = False
    Exit Function

TestFail:
    If Err.Number = ExpectedError Then
        Test_DictsToTable_fail_1 = True
        Exit Function
    Else
        Test_DictsToTable_fail_1 = False
        Exit Function
    End If
End Function


Function Test_DictsToTable_fail_2(TableDict As Collection, RangeObj As Range)
    Const ExpectedError As Long = -996
    On Error GoTo TestFail

    DictsToTable TableDict, RangeObj, "someName"
    
    Test_DictsToTable_fail_2 = True
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_DictsToTable_fail_2 = True
        Exit Function
    Else
        Test_DictsToTable_fail_2 = True
        Exit Function
    End If
End Function
