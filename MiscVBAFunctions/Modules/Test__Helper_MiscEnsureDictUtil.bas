Attribute VB_Name = "Test__Helper_MiscEnsureDictUtil"
Option Explicit


Function Test_EnsureDictIContainer(C As Collection)
    Dim Pass As Boolean
    Pass = True

    Pass = False = C(1).Exists("a") = Pass = True
    Pass = True = C(2).Exists("b") = Pass = True
    Pass = False = C(3).Exists("c") = Pass = True
    
    Pass = True = C(1).Exists("A") = Pass = True
    Pass = False = C(2).Exists("B") = Pass = True
    Pass = True = C(3).Exists("C") = Pass = True
    
    Test_EnsureDictIContainer = Pass
End Function


Function Test_EnsureDictIContainer_I(C As Collection)
    Dim Pass As Boolean
    Pass = True

    Pass = True = C(1).Exists("a") = Pass = True
    Pass = True = C(2).Exists("b") = Pass = True
    Pass = True = C(3).Exists("c") = Pass = True
    
    Pass = True = C(1).Exists("A") = Pass = True
    Pass = True = C(2).Exists("B") = Pass = True
    Pass = True = C(3).Exists("C") = Pass = True
    
    Test_EnsureDictIContainer_I = Pass
End Function

