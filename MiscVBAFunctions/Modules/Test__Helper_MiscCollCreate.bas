Attribute VB_Name = "Test__Helper_MiscCollCreate"
Option Explicit

Function Test_zip(C1, C2)
    Dim Pass As Boolean
    Pass = True

    Dim Cout As Collection
    Set Cout = zip(C1, C2)

    Pass = 1 = Cout(1)(1) = Pass
    Pass = 4 = Cout(1)(2) = Pass
    Pass = 2 = Cout(2)(1) = Pass
    Pass = 5 = Cout(2)(2) = Pass
    Pass = 3 = Cout(3)(1) = Pass
    Pass = 6 = Cout(3)(2) = Pass
    
    Test_zip = Pass
End Function
