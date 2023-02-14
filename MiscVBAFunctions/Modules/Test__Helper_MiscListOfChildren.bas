Attribute VB_Name = "Test__Helper_MiscListOfChildren"
Option Explicit

Function Test_GetListOfChildren()
    Dim Pass As Boolean
    Pass = True
    
    Dim Depths As Collection
    Set Depths = Col( _
    1, _
      2, _
      2, _
        3, _
        3, _
          4, _
      2, _
    1, _
    1, _
      2)
    
    Dim ChildrenDepths As Collection
    Set ChildrenDepths = GetListOfChildren(Depths)
    
    Pass = ChildrenDepths(1)(1) = CLng(2) = Pass
    Pass = ChildrenDepths(1)(2) = CLng(3) = Pass
    Pass = ChildrenDepths(1)(3) = CLng(7) = Pass
    
    Pass = ChildrenDepths(2).Count = CLng(0) = Pass
    
    Pass = ChildrenDepths(3)(1) = CLng(4) = Pass
    Pass = ChildrenDepths(3)(2) = CLng(5) = Pass
    
    Pass = ChildrenDepths(4).Count = CLng(0) = Pass
    
    Pass = ChildrenDepths(5)(1) = CLng(6) = Pass
    
    Pass = ChildrenDepths(6).Count = CLng(0) = Pass
    Pass = ChildrenDepths(7).Count = CLng(0) = Pass
    Pass = ChildrenDepths(8).Count = CLng(0) = Pass
    
    Pass = ChildrenDepths(9)(1) = CLng(10) = Pass
    
    ' test back / upwards children
    Set ChildrenDepths = GetListOfChildren(Depths, False)
    Pass = ChildrenDepths(7)(1) = CLng(5) = Pass
    Pass = ChildrenDepths(7)(2) = CLng(4) = Pass
    
    Pass = ChildrenDepths(8)(1) = CLng(7) = Pass
    Pass = ChildrenDepths(8)(2) = CLng(3) = Pass
    Pass = ChildrenDepths(8)(3) = CLng(2) = Pass
    
    Test_GetListOfChildren = Pass
End Function

