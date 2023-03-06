Attribute VB_Name = "Test__Helper_MiscArray"
Option Explicit

Function Test_ErrorToNullStringTransformation_1()
    Dim Arr(2, 2)
    Dim Pass As Boolean
    Pass = True
    
    Arr(0, 0) = 100.2: Arr(0, 1) = CVErr(XlErrName)
    Arr(1, 0) = 2.1: Arr(1, 1) = CVErr(XlErrNA)
    ErrorToNullStringTransformation Arr

    Pass = 100.2 = Arr(0, 0) = Pass = True
    Pass = VbNullString = Arr(0, 1) = Pass = True
    Pass = 2.1 = Arr(1, 0) = Pass = True
    Pass = VbNullString = Arr(1, 1) = Pass = True
    
    Test_ErrorToNullStringTransformation_1 = Pass
End Function


Function Test_ErrorToNullStringTransformation_2()
    Dim Arr(3)
    Dim Pass As Boolean
    Pass = True
    
    Arr(0) = 1.2: Arr(1) = CVErr(XlErrRef): Arr(2) = 3.8
    ErrorToNullStringTransformation Arr

    Pass = 1.2 = Arr(0) = Pass = True
    Pass = VbNullString = Arr(1) = Pass = True
    Pass = 3.8 = Arr(2) = Pass = True
    
    Test_ErrorToNullStringTransformation_2 = Pass
End Function


Function Test_ArrayToRange_fail(Arr() As Variant, RangeObj As Range)
    Const ExpectedError As Long = 9
    On Error GoTo TestFail

    Dim LO As ListObject
    Set LO = ArrayToRange(Arr, RangeObj, True)
    
    Test_ArrayToRange_fail = False
    Exit Function

TestFail:
    If Err.Number = ExpectedError Then
        Test_ArrayToRange_fail = True
        Exit Function
    Else
        Test_ArrayToRange_fail = False
        Exit Function
    End If
End Function


Function Test_ArrayToNewTable_fail(Arr() As Variant, RangeObj As Range)
    Const ExpectedError As Long = -999
    On Error GoTo TestFail
    
    Dim LO As ListObject
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, True)
    Set LO = ArrayToNewTable("TestTable", Arr, RangeObj, True)

    Test_ArrayToNewTable_fail = False
    Exit Function
        
TestFail:
    If Err.Number = ExpectedError Then
        Test_ArrayToNewTable_fail = True
        Exit Function
    Else
        Test_ArrayToNewTable_fail = False
        Exit Function
    End If
End Function
