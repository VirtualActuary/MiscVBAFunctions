Attribute VB_Name = "Test__MiscDataContainers"
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

'@TestMethod("MiscDataStructures")
Private Sub Test_Join_Collections()
    On Error GoTo TestFail

    'Arrange:
    Dim z, x, y As New Collection
    
    'Act:
    Set x = col(1, 2, 3)
    Set y = col(4, 5, 6)
    Set z = JoinContainers(x, y)
    
    'Assert:
    Assert.AreEqual 1, x(1)
    Assert.AreEqual 2, x(2)
    Assert.AreEqual 3, x(3)
    Assert.AreEqual 4, y(1)
    Assert.AreEqual 5, y(2)
    Assert.AreEqual 6, y(3)

    Assert.AreEqual 1, z(1)
    Assert.AreEqual 2, z(2)
    Assert.AreEqual 3, z(3)
    Assert.AreEqual 4, z(4)
    Assert.AreEqual 5, z(5)
    Assert.AreEqual 6, z(6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDataStructures")
Private Sub Test_Join_Collections_fail()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim z, x, y As New Collection
    
    'Act:
    Set x = col(1, 2, 3)
    Set y = col(4, 5, 6)
    Set z = JoinContainers(x, y)
    
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

'@TestMethod("MiscDataStructures")
Private Sub Test_Join_Collections_fail_2()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim d, d1 As Dictionary
    Dim c As New Collection
    'Act:

    Set d1 = dict("a", 1, "b", 2)
    Set c = col(1, 2, 3)
    
    Set d = JoinContainers(d1, c)
    

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

'@TestMethod("MiscDataStructures")
Private Sub Test_Concat_Collections()
    On Error GoTo TestFail
    
    Dim x, y As New Collection
    
    'Act:
    Set x = col(1, 2, 3)
    Set y = col(4, 5, 6)
    ConcatContainers x, y
    
    'Assert:

    Assert.AreEqual 4, y(1)
    Assert.AreEqual 5, y(2)
    Assert.AreEqual 6, y(3)

    Assert.AreEqual 1, x(1)
    Assert.AreEqual 2, x(2)
    Assert.AreEqual 3, x(3)
    Assert.AreEqual 4, x(4)
    Assert.AreEqual 5, x(5)
    Assert.AreEqual 6, x(6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDataStructures")
Private Sub Test_Concat_Collections_fail()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim d, d1 As Dictionary
    Dim c As New Collection
    'Act:

    Set d1 = dict("a", 1, "b", 2)
    Set c = col(1, 2, 3)
    
    ConcatContainers d1, c
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

'@TestMethod("MiscDataStructures")
Private Sub Test_Join_Dicts()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d, d1, d2 As Dictionary

    'Act:
    Set d1 = dict("a", 1, "b", 2)
    Set d2 = dict("c", 10, "d", 20)
    Set d = JoinContainers(d1, d2)
    
    'Assert:
    Assert.AreEqual 1, d("a")
    Assert.AreEqual 2, d("b")
    Assert.AreEqual 10, d("c")
    Assert.AreEqual 20, d("d")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscDataStructures")
Private Sub Test_Concat_Dicts()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d1, d2 As Dictionary

    'Act:
    Set d1 = dict("a", 1, "b", 2)
    Set d2 = dict("c", 10, "d", 20)
    ConcatContainers d1, d2
    
    'Assert:
    Assert.AreEqual 1, d1("a")
    Assert.AreEqual 2, d1("b")
    Assert.AreEqual 10, d1("c")
    Assert.AreEqual 20, d1("d")
    Assert.AreEqual 10, d2("c")
    Assert.AreEqual 20, d2("d")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
