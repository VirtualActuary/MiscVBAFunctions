Attribute VB_Name = "Test__Helper_MiscHasKey"
Option Explicit

Function Test_HasKey_Collection()
    Dim C As New Collection
    Dim Pass As Boolean
    Pass = True
    
    C.Add "foo", "a"
    C.Add Col("x", "y", "z"), "b"
    
    Pass = True = HasKey(C, "a") = Pass = True ' True for scalar
    Pass = True = HasKey(C, "b") = Pass = True ' True for scalar
    Pass = True = HasKey(C, "A") = Pass = True ' True for case insensitive
    
    Test_HasKey_Collection = Pass
End Function


Function Test_HasKey_Workbook()
    Test_HasKey_Workbook = True = HasKey(Workbooks, ThisWorkbook.Name)
End Function


Function Test_HasKey_Dictionary_object()
    Dim Pass As Boolean
    Pass = True
    
    Dim DObj As Object
    Set DObj = CreateObject("Scripting.Dictionary")
    
    DObj.Add "a", "foo"
    DObj.Add "b", Col("x", "y", "z")
    
    Pass = HasKey(DObj, "a") = Pass = True ' True for scalar
    Pass = HasKey(DObj, "b") = Pass = True ' True for scalar
    Pass = (False = HasKey(DObj, "A")) = Pass = True ' False - case sensitive by default

    Test_HasKey_Dictionary_object = Pass
End Function


Function Test_HasKey_Dictionary_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail

    HasKey 5, "a"
    HasKey ThisWorkbook, "A"

    Test_HasKey_Dictionary_fail = False
    Exit Function
    
TestFail:
    If Err.Number = ExpectedError Then
        Test_HasKey_Dictionary_fail = True
        Exit Function
    Else
        Test_HasKey_Dictionary_fail = False
        Exit Function
    End If
End Function
