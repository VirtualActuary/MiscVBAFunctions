Attribute VB_Name = "Test__Helper_MiscCollection"
Option Explicit
 
Function Test_min_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail

    Dim C As Collection
    min C
    
    Test_min_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_min_fail = True
        Exit Function
    Else
        Test_min_fail = False
        Exit Function
    End If
End Function


Function Test_max_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail

    Dim C As Collection
    max C
    
    Test_max_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_max_fail = True
        Exit Function
    Else
        Test_max_fail = False
        Exit Function
    End If
End Function


Function Test_mean_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail

    Dim C As Collection
    mean C
    
    Test_mean_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_mean_fail = True
        Exit Function
    Else
        Test_mean_fail = False
        Exit Function
    End If
End Function


Function Test_Join_Collections_fail_2(D1 As Dictionary, C As Collection)
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    Dim D As Dictionary
    Set D = JoinCollections(D1, C)
    
    Test_Join_Collections_fail_2 = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_Join_Collections_fail_2 = True
        Exit Function
    Else
        Test_Join_Collections_fail_2 = False
        Exit Function
    End If
End Function


Function Test_Concat_Collections_fail(D1 As Dictionary, C As Collection)
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    ConcatCollections D1, C
    
    Test_Concat_Collections_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_Concat_Collections_fail = True
        Exit Function
    Else
        Test_Concat_Collections_fail = False
        Exit Function
    End If
End Function


Function Test_Concat_Collections(X, Y, Z)
    Dim Pass As Boolean
    Pass = True

    ConcatCollections X, Z, Y

    Pass = 1 = X(1) = Pass
    Pass = 2 = X(2) = Pass
    Pass = 5 = X(3) = Pass
    Pass = 6 = X(4) = Pass
    Pass = 3 = X(5) = Pass
    Pass = 4 = X(6) = Pass

    Test_Concat_Collections = Pass
End Function


Function Test_CollectionToArray_empty(Arr() As Variant)
    Dim Pass As Boolean
    Pass = True

    Pass = IsArray(Arr) = Pass
    Pass = CLng(0) = LBound(Arr) = Pass
    Pass = CLng(-1) = UBound(Arr) = Pass
    
    Test_CollectionToArray_empty = Pass
End Function
