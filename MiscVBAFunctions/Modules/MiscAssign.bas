Attribute VB_Name = "MiscAssign"
Option Explicit

Public Function assign(ByRef var As Variant, ByRef val As Variant)
    ' Assign a value to a variable and also return that value. The goal of this function is to
    ' overcome the different `set` syntax for assigning an object vs. assigning a native type
    ' like Int, Double etc. Additionally this function has similar functionality to Python's
    ' walrus operator: https://towardsdatascience.com/the-walrus-operator-7971cd339d7d
    '
    ' Args:
    '   var: The input variable that could be an object.
    '   val: The value that the var input needs to be changed to.
    '
    ' Returns:
    '   The value from the input.
    
    If IsObject(val) Then 'Object
        Set var = val
        Set assign = val
    Else 'Variant
        var = val
        assign = val
    End If
End Function
