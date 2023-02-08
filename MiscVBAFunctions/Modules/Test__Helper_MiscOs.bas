Attribute VB_Name = "Test__Helper_MiscOs"
Option Explicit

Function Test_is64BitXl()
    #If Win64 Then
        Test_is64BitXl = Is64BitXl()
    #Else
        Test_is64BitXl = Not Is64BitXl()
    #End If
End Function


