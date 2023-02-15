Attribute VB_Name = "Test__Helper_MiscRemvoeGridline"
Option Explicit

Function TestRemoveGridlines()
    RemoveGridLines ThisWorkbook.Sheets(1)
    TestRemoveGridlines = Not ThisWorkbook.Windows(1).SheetViews(1).DisplayGridlines
End Function

