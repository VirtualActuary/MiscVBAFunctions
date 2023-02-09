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
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 4, min(col(7, 4, 5, 6)), "min test succeeded"
    Assert.AreEqual 5, min(col(9, 5, 6)), "min test succeeded"

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
    
    
    min C
    
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
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 6, max(col(4, 5, 6, 1, 2)), "max test succeeded"
    Assert.AreEqual 6.1, max(col(5.3, 6.1)), "max test succeeded"


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
    max C

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
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual 4#, mean(col(4, 5, 6, 3, 2)), "mean test succeeded"
    Assert.AreEqual 6#, mean(col(5, 7)), "mean test succeeded"

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
    mean C

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
    
    'Arrange:

    'Act:

    'Assert:

    Assert.IsTrue IsValueInCollection(col("a", "b"), "b")
    Assert.IsFalse IsValueInCollection(col("a", "b"), "c")
    Assert.IsFalse IsValueInCollection(col("a", "b"), "B", True)

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
    Dim X As New Collection
    Dim Y As New Collection
    Dim Z As New Collection
    
    'Act:
    Set w = col(1, 2)
    Set X = col(3, 4)
    Set Y = col(5, 6)
    Set Z = JoinCollections(X, Y, w)
    
    'Assert:
    Assert.AreEqual 3, Z(1)
    Assert.AreEqual 4, Z(2)
    Assert.AreEqual 5, Z(3)
    Assert.AreEqual 6, Z(4)
    Assert.AreEqual 1, Z(5)
    Assert.AreEqual 2, Z(6)

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
    Dim Z As New Collection
    Dim X As New Collection
    Dim Y As New Collection
    
    'Act:
    Set X = col(1, 2, 3)
    Set Y = col(4, 5, 6)
    Set Z = JoinCollections(X, Y)
    
    'Assert:
    Debug.Print X(4)
    Debug.Print Y(4)
    Debug.Print Z(7)

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
    Set C = col(1, 2, 3)
    
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

    Dim X As Collection
    Dim Y As Collection
    Dim Z As Collection

    'Act:
    Set X = col(1, 2)
    Set Y = col(3, 4)
    Set Z = col(5, 6)
    ConcatCollections X, Z, Y
    
    'Assert:

    Assert.AreEqual 1, X(1)
    Assert.AreEqual 2, X(2)
    Assert.AreEqual 5, X(3)
    Assert.AreEqual 6, X(4)
    Assert.AreEqual 3, X(5)
    Assert.AreEqual 4, X(6)

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
    Set C = col(1, 2, 3)
    
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
    Set C1 = col(7, 4, 5, 6)

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
    Set C1 = col()

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
