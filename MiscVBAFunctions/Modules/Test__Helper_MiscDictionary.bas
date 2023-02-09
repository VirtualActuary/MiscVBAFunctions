Attribute VB_Name = "Test__Helper_MiscDictionary"
Option Explicit

Function Test_dictget_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail

    Dim D As Dictionary
    Set D = dict("a", 2, "b", ThisWorkbook)
    dictget D, "c"
    
    Test_dictget_fail = False
    Exit Function

TestFail:
    If Err.Number = ExpectedError Then
        Test_dictget_fail = True
        Exit Function
    Else
        Test_dictget_fail = False
        Exit Function
    End If
End Function
