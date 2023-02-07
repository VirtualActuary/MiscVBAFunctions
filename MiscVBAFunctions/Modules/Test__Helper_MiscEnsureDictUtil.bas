Attribute VB_Name = "Test__Helper_MiscEnsureDictUtil"
Option Explicit


Function Test_EnsureDictIContainer(C As Collection)
    Dim Pass As Boolean
    Pass = True

    Pass = False = C(1).Exists("a") = Pass
    Pass = True = C(2).Exists("b") = Pass
    Pass = False = C(3).Exists("c") = Pass
    
    Pass = True = C(1).Exists("A") = Pass
    Pass = False = C(2).Exists("B") = Pass
    Pass = True = C(3).Exists("C") = Pass
    
    Test_EnsureDictIContainer = Pass
End Function


Function Test_EnsureDictIContainer_I(C As Collection)
    Dim Pass As Boolean
    Pass = True

    Pass = True = C(1).Exists("a") = Pass
    Pass = True = C(2).Exists("b") = Pass
    Pass = True = C(3).Exists("c") = Pass
    
    Pass = True = C(1).Exists("A") = Pass
    Pass = True = C(2).Exists("B") = Pass
    Pass = True = C(3).Exists("C") = Pass
    
    Test_EnsureDictIContainer_I = Pass
End Function

