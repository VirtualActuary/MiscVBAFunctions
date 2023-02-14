Attribute VB_Name = "Test__Helper_MiscCollection"
Option Explicit
 
Function Test_min_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail

    Dim Col1 As Collection
    Min Col1
    
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

    Dim Col1 As Collection
    Max Col1
    
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

    Dim Col1 As Collection
    Mean Col1
    
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


Function Test_Join_Collections_fail_2(D1 As Dictionary, C1 As Collection)
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    Dim DOut As Dictionary
    Set DOut = JoinCollections(D1, C1)
    
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


Function Test_Concat_Collections_fail(D1 As Dictionary, C1 As Collection)
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    ConcatCollections D1, C1
    
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


Function Test_Concat_Collections(C1, C2, C3)
    Dim Pass As Boolean
    Pass = True

    ConcatCollections C1, C2, C3

    Pass = 1 = C1(1) = Pass
    Pass = 2 = C1(2) = Pass
    Pass = 5 = C1(3) = Pass
    Pass = 6 = C1(4) = Pass
    Pass = 3 = C1(5) = Pass
    Pass = 4 = C1(6) = Pass

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
