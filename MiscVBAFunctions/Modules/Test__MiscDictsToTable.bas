Attribute VB_Name = "Test__MiscDictsToTable"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("MiscDictToTable")
Private Sub Test_DictsToTable_1()
    On Error GoTo TestFail
    
    'Arrange:

    Dim WB As Workbook
    Set WB = ExcelBook("")

    Dim TableDict As Collection
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim Table As ListObject
    

    'Act:
    With dict
        .Add "col1", 1
        .Add "col2", 2
        .Add "col3", 3
    End With
    
    Set TableDict = col(dict, dict)
    
    Set Table = DictsToTable(TableDict, WB.Worksheets(1).Range("A1"), "someName")
    
    'Assert:
    Assert.AreEqual "col1", Table.ListColumns.Item(1).Name
    Assert.AreEqual "col2", Table.ListColumns.Item(2).Name
    
    Assert.AreEqual 1, CInt(Table.ListRows.Item(1).Range(1))
    Assert.AreEqual 2, CInt(Table.ListRows.Item(1).Range(2))
    Assert.AreEqual 3, CInt(Table.ListRows.Item(1).Range(3))
    
    Assert.AreEqual 1, CInt(Table.ListRows.Item(2).Range(1))
    Assert.AreEqual 2, CInt(Table.ListRows.Item(2).Range(2))
    Assert.AreEqual 3, CInt(Table.ListRows.Item(2).Range(3))
TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDictToTable")
Private Sub Test_DictsToTable_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook("")
    Dim TableDict As Collection
    Dim dict1 As Dictionary
    Set dict1 = New Dictionary
    Dim dict2 As Dictionary
    Set dict2 = New Dictionary
    Dim Table As ListObject

    'Act:
    With dict1
        .Add "col1", 1
        .Add "col2", 2
        .Add "col3", 3
    End With
    
    With dict2
        .Add "col1", 10
        .Add "col3", 20
        .Add "col2", 30
    End With
    
    Set TableDict = col(dict1, dict2)
    
    Set Table = DictsToTable(TableDict, WB.Worksheets(1).Range("A5"), "someName")
    
    'Assert:
    Assert.AreEqual "col1", Table.ListColumns.Item(1).Name
    Assert.AreEqual "col2", Table.ListColumns.Item(2).Name
    Assert.AreEqual "col3", Table.ListColumns.Item(3).Name
    
    Assert.AreEqual 1, CInt(Table.ListRows.Item(1).Range(1))
    Assert.AreEqual 2, CInt(Table.ListRows.Item(1).Range(2))
    Assert.AreEqual 3, CInt(Table.ListRows.Item(1).Range(3))
    
    Assert.AreEqual 10, CInt(Table.ListRows.Item(2).Range(1))
    Assert.AreEqual 30, CInt(Table.ListRows.Item(2).Range(2))  ' Note pos 2 and 3 switched.
    Assert.AreEqual 20, CInt(Table.ListRows.Item(2).Range(3))

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDictToTable")
Private Sub Test_DictsToTable_fail_1()
    Const ExpectedError As Long = -997
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook("")
    Dim TableDict As Collection
    Dim dict1 As Dictionary
    Set dict1 = New Dictionary
    Dim dict2 As Dictionary
    Set dict2 = New Dictionary
    Dim Table As ListObject

    'Act:
    With dict1
        .Add "col1", 1
        .Add "col2", 2
        .Add "col3", 3
        .Add "col4", 30
    End With
    
    With dict2
        .Add "col1", 10
        .Add "col2", 20
        .Add "col3", 30
        
    End With
    
    Set TableDict = col(dict1, dict2)
    
    DictsToTable TableDict, WB.Worksheets(1).Range("A10"), "someName"

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscDictToTable")
Private Sub Test_DictsToTable_fail_2()
    Const ExpectedError As Long = -996
    On Error GoTo TestFail
    
    'Arrange:
    Dim WB As Workbook
    Set WB = ExcelBook("")
    Dim TableDict As Collection
    Dim dict1 As Dictionary
    Set dict1 = New Dictionary
    Dim dict2 As Dictionary
    Set dict2 = New Dictionary
    Dim Table As ListObject

    'Act:
    With dict1
        .Add "col1", 1
        .Add "col2", 2
        .Add "col3", 3
    End With
    
    With dict2
        .Add "col1", 10
        .Add "col2", 20
        .Add "col4", 30
        
    End With
    
    Set TableDict = col(dict1, dict2)
    
    DictsToTable TableDict, WB.Worksheets(1).Range("A15"), "someName"

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    WB.Close False
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
