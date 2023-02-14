Attribute VB_Name = "Test__Helper_MiscCollCreate"
Option Explicit

Function Test_zip(C1, C2)
    Dim Pass As Boolean
    Pass = True

    Dim Cout As Collection
    Set Cout = Zip(C1, C2)

    Pass = 1 = Cout(1)(1) = Pass = True
    Pass = 4 = Cout(1)(2) = Pass = True
    Pass = 2 = Cout(2)(1) = Pass = True
    Pass = 5 = Cout(2)(2) = Pass = True
    Pass = 3 = Cout(3)(1) = Pass = True
    Pass = 6 = Cout(3)(2) = Pass = True
    
    Test_zip = Pass
End Function
