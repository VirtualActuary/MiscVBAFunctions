Attribute VB_Name = "Test__MiscCollection"
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

'@TestMethod("MiscCollection.min")
Private Sub Test_min()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual 4, Min(Col(7, 4, 5, 6)), "min test succeeded"
    Assert.AreEqual 5, Min(Col(9, 5, 6)), "min test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.min")
Private Sub Test_min_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    'Act:
       
    Min C
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Assert.Succeed
        
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection.max")
Private Sub Test_max()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual 6, Max(Col(4, 5, 6, 1, 2)), "max test succeeded"
    Assert.AreEqual 6.1, Max(Col(5.3, 6.1)), "max test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.max")
Private Sub Test_max_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection

    'Act:
    Max C

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection.mean")
Private Sub Test_mean()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual 4#, Mean(Col(4, 5, 6, 3, 2)), "mean test succeeded"
    Assert.AreEqual 6#, Mean(Col(5, 7)), "mean test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection.mean")
Private Sub Test_mean_fail()
    Const ExpectedError As Long = 91
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection

    'Act:
    Mean C

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection.IsValueInCollection")
Private Sub Test_IsValueInCollection()
    On Error GoTo TestFail

    'Assert:
    Assert.IsTrue IsValueInCollection(Col("a", "b"), "b")
    Assert.IsFalse IsValueInCollection(Col("a", "b"), "c")
    Assert.IsFalse IsValueInCollection(Col("a", "b"), "B", True)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_Join_Collections()
    On Error GoTo TestFail

    'Arrange:
    Dim w As New Collection
    Dim x As New Collection
    Dim y As New Collection
    Dim z As New Collection
    
    'Act:
    Set w = Col(1, 2)
    Set x = Col(3, 4)
    Set y = Col(5, 6)
    Set z = JoinCollections(x, y, w)
    
    'Assert:
    Assert.AreEqual 3, z(1)
    Assert.AreEqual 4, z(2)
    Assert.AreEqual 5, z(3)
    Assert.AreEqual 6, z(4)
    Assert.AreEqual 1, z(5)
    Assert.AreEqual 2, z(6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_Join_Collections_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim z As New Collection
    Dim x As New Collection
    Dim y As New Collection
    
    'Act:
    Set x = Col(1, 2, 3)
    Set y = Col(4, 5, 6)
    Set z = JoinCollections(x, y)
    
    'Assert:
    Debug.Print x(4)
    Debug.Print y(4)
    Debug.Print z(7)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_Join_Collections_fail_2()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim D As Dictionary
    Dim D1 As Dictionary
    Dim C As New Collection
    'Act:

    Set D1 = Dict("a", 1, "b", 2)
    Set C = Col(1, 2, 3)
    
    Set D = JoinCollections(D1, C)
    

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_Concat_Collections()
    On Error GoTo TestFail

    Dim x As Collection
    Dim y As Collection
    Dim z As Collection

    'Act:
    Set x = Col(1, 2)
    Set y = Col(3, 4)
    Set z = Col(5, 6)
    ConcatCollections x, z, y
    
    'Assert:

    Assert.AreEqual 1, x(1)
    Assert.AreEqual 2, x(2)
    Assert.AreEqual 5, x(3)
    Assert.AreEqual 6, x(4)
    Assert.AreEqual 3, x(5)
    Assert.AreEqual 4, x(6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_Concat_Collections_fail()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim D As Dictionary
    Dim D1 As Dictionary
    Dim C As New Collection
    'Act:

    Set D1 = Dict("a", 1, "b", 2)
    Set C = Col(1, 2, 3)
    
    ConcatCollections D1, C
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub



'@TestMethod("MiscCollection.CollectionToArray")
Private Sub Test_CollectionToArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C1 As Collection
    Set C1 = Col(7, 4, 5, 6)

    'Act:
    Dim a1 As Variant
    a1 = CollectionToArray(C1)

    'Assert:
    
    Assert.IsTrue IsArray(a1), "Result is an array"
    
    Dim expectedLowerBound As Long
    expectedLowerBound = 0
    Assert.AreEqual expectedLowerBound, LBound(a1), "Lower bound"
    
    Dim expectedUpperBound As Long
    expectedUpperBound = 3
    Assert.AreEqual expectedUpperBound, UBound(a1), "Upper bound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("MiscCollection.CollectionToArray")
Private Sub Test_CollectionToArray_empty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C1 As Collection
    Set C1 = Col()

    'Act:
    Dim a1 As Variant
    a1 = CollectionToArray(C1)

    'Assert:
    
    Assert.IsTrue IsArray(a1), "Result is an array"
    
    Dim expectedLowerBound As Long
    expectedLowerBound = 0
    Assert.AreEqual expectedLowerBound, LBound(a1), "Lower bound"
    
    Dim expectedUpperBound As Long
    expectedUpperBound = -1
    Assert.AreEqual expectedUpperBound, UBound(a1), "Upper bound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_indexOf()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Collection
    Dim C2 As Collection

    'Act:
    Set C = Col("variables10", 0, "variables", 10, "variables2", "20", "variables_10", 30, "variables_2", 40)
    Set C2 = Col(12, 23, 34, 45, 56, 67)
    
    'Assert:
    Assert.AreEqual 3, CInt(IndexOf(C, "variables"))
    Assert.AreEqual 2, CInt(IndexOf(C, 0))
    Assert.AreEqual 0, CInt(IndexOf(C, "Foo"))

    Assert.AreEqual 5, CInt(IndexOf(C2, 56))
    Assert.AreEqual 0, CInt(IndexOf(C2, "23"))

    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscCollection")
Private Sub Test_uniqueCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Col1 As Collection
    Dim Col2 As Collection

    'Act:
    Set Col1 = Col("1", "3.4", 3.4, 1)
    Set Col2 = Col(3.4, 3.4, "1", "asdf")

    'Assert:
    
    Assert.AreEqual CLng(4), UniqueCollection(Col1).Count
    Assert.AreEqual CLng(3), UniqueCollection(Col2).Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
