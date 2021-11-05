Attribute VB_Name = "MiscF"
Option Explicit

'************"Casing"
' Uncomment and comment block to get casing back for the project


'Dim J
'Dim I
'Dim WB
'Dim WS

'************"MiscAssign"
' Assign a value to a variable and also return that value. The goal of this function is to
' overcome the different `set` syntax for assigning an object vs. assigning a native type
' like Int, Double etc. Additionally this function has similar functionality to Python's
' walrus operator: https://towardsdatascience.com/the-walrus-operator-7971cd339d7d

Function assign(ByRef var As Variant, ByRef val As Variant)
    If IsObject(val) Then 'Object
        Set var = val
        Set assign = val
    Else 'Variant
        var = val
        assign = val
    End If
End Function

'************"MiscCollection"



Function min(ByVal col As Collection) As Variant
    
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    min = col(1)
    
    For Each Entry In col
        If Entry < min Then
            min = Entry
        End If
    Next Entry
    
    
    
End Function

Function max(ByVal col As Collection) As Variant
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    max = col(1)
    Dim Entry As Variant
    
    For Each Entry In col
        If Entry > max Then
            max = Entry
        End If
    Next Entry

End Function

Function mean(ByVal col As Collection) As Variant
    If col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If

    mean = 0
    Dim Entry As Variant
    
    For Each Entry In col
        mean = mean + Entry
    Next Entry
    
    mean = mean / col.Count
    
End Function






'************"MiscCollectionCreate"


Function col(ParamArray Args() As Variant) As Collection
    Set col = New Collection
    Dim i As Long

    For i = LBound(Args) To UBound(Args)
        col.Add Args(i)
    Next

End Function


Function zip(ParamArray Args() As Variant) As Collection
    Dim i As Long
    Dim J As Long
    
    Dim N As Long
    Dim M As Long
    

    M = -1
    For i = LBound(Args) To UBound(Args)
        If M = -1 Then
            M = Args(i).Count
        ElseIf Args(i).Count < M Then
            M = Args(i).Count
        End If
    Next i

    Set zip = New Collection
    Dim ICol As Collection
    For i = 1 To M
        Set ICol = New Collection
        For J = LBound(Args) To UBound(Args)
            ICol.Add Args(J).Item(i)
        Next J
        zip.Add ICol
    Next i
End Function




'************"MiscCreateTextFile"


Private Sub testCreateTextFile()
    CreateTextFile "foo", ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt"
    ' TODO: assertion
End Sub

Function CreateTextFile(Content As String, FilePath As String)
    ' Creates a new / overwrites an existing text file with Content
    
    Dim oFile As Integer
    oFile = FreeFile
    
    Open FilePath For Output As #oFile
        Print #oFile, Content
    Close #oFile

End Function

'************"MiscDictionary"


Private Function testDictget()

    Dim d As Dictionary
    Set d = dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(d, "a"), 2 ' returns 2
    Debug.Print dictget(d, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(d, "c", vbNullString), vbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(d, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Function


Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
        
    If d.Exists(key) Then
        assign dictget, d.Item(key)
        
    ElseIf Not IsEmpty(default) Then
        assign dictget, default
        
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & key & "` "
            errmsg = errmsg & "not in dictionary"
        On Error GoTo 0
        
        Err.Raise 9, , errmsg
    End If
End Function

'************"MiscDictionaryCreate"


Function dict(ParamArray Args() As Variant) As Dictionary
    'Case sensitive dictionary
    
    Dim errmsg As String
    Set dict = New Dictionary
    
    Dim i As Long
    Dim Cnt As Long
    Cnt = 0
    For i = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If i + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(i) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dict.Add Args(i), Args(i + 1)
Cont:
    Next i

End Function


Function dicti(ParamArray Args() As Variant) As Dictionary
    'Case insensitive dictionary
    
    Dim errmsg As String
    Set dicti = New Dictionary
    dicti.CompareMode = TextCompare
    
    Dim i As Long
    Dim Cnt As Long
    Cnt = 0
    For i = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If i + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(i) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dicti.Add Args(i), Args(i + 1)
Cont:
    Next i

End Function


'************"MiscFreezePanes"



Private Sub test()
    
    On Error GoTo UnFreeze
    
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Sheets("Sheet1")
    FreezePanes WS.Range("D4")
    
UnFreeze:
    UnFreezePanes WS
    
End Sub

Sub FreezePanes(r As Range)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    Dim WS As Worksheet
    Set WS = r.Parent
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With Application.Windows(WS.Parent.Name)
        ' if existing freezed panes, remove them
        If .FreezePanes = True Then
            .FreezePanes = False
        End If
        Application.GoTo WS.Cells(1, 1) ' <- to ensure we don't hide the top/ left side of sheet
        ' Unfortunately, we have to do this :/
        Application.GoTo r
        .FreezePanes = True
    End With
    
    Application.ScreenUpdating = currentScreenUpdating
    
    CurrentActiveSheet.Activate
End Sub

Sub UnFreezePanes(WS As Worksheet)
    
    Dim CurrentActiveSheet As Worksheet
    Set CurrentActiveSheet = ActiveSheet
    
    ' Unfortunately, we have to do this :/
    WS.Activate
    With Application.Windows(WS.Parent.Name)
        .FreezePanes = False
    End With
    
    CurrentActiveSheet.Activate
End Sub

'************"MiscGetUniqueItems"


Private Function TestGetUniqueItems()
    Dim arr(3)
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "b"
    Debug.Print UBound(GetUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(arr), 1), 3 ' zero index + case sensitive
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(GetUniqueItems(arr, False), 1), 2 ' zero index + case insensitive
    
    arr(0) = 1: arr(1) = 2: arr(2) = 3: arr(3) = 2
    Debug.Print UBound(GetUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = 1: arr(1) = 1: arr(2) = "a": arr(3) = "a"
    Debug.Print UBound(GetUniqueItems(arr), 1), 1 ' zero index
    
End Function

Function GetUniqueItems(arr() As Variant, _
            Optional CaseSensitive As Boolean = True) As Variant
    If ArrayLen(arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim d As New Dictionary
        If Not CaseSensitive Then
            d.CompareMode = TextCompare
        End If
        
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            If Not d.Exists(arr(i)) Then
                d.Add arr(i), arr(i)
            End If
        Next
        
        GetUniqueItems = d.Keys()
    End If
End Function


' Returns the number of elements in an array for a given dimension.
Private Function ArrayLen(arr As Variant, _
    Optional dimNum As Integer = 1) As Long
    
    If IsEmpty(arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function


'************"MiscGroupOnIndentations"


Private Sub TestGroupOnIndentations()

    ' test rows
    GroupRowsOnIndentations ThisWorkbook.Names("__TestGroupRowsOnIndentations__").RefersToRange
    ' test columns
    GroupColumnsOnIndentations ThisWorkbook.Names("__TestGroupColumnsOnIndentations__").RefersToRange

End Sub

Sub GroupRowsOnIndentations(r As Range)
    ' groups the rows based on indentations of the cells in the range
    
    Dim ri As Range, WS As Worksheet
    For Each ri In r
        ri.EntireRow.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub


Sub GroupColumnsOnIndentations(r As Range)
    ' groups the columns based on indentations of the cells in the range
    
    Dim ri As Range, WS As Worksheet
    For Each ri In r
        ri.EntireColumn.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub


Private Sub TestRemoveGroupings()
    ' Test rows
    RemoveRowGroupings ThisWorkbook.Sheets("GroupOnIndentations")
    ' Test columns
    RemoveColumnGroupings ThisWorkbook.Sheets("GroupOnIndentations")
End Sub


Sub RemoveRowGroupings(WS As Worksheet)
    Dim r As Range, ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    For Each ri In r.Columns(1)
        ri.EntireRow.OutlineLevel = 1
    Next ri
End Sub

Sub RemoveColumnGroupings(WS As Worksheet)
    Dim r As Range, ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    For Each ri In r.Rows(1)
        ri.EntireColumn.OutlineLevel = 1
    Next ri
End Sub


'************"MiscHasKey"


Private Sub TestHasKey()

    Dim c As New Collection
    c.Add "a", "a"
    c.Add col("x", "y", "z"), "b"
    
    Debug.Print vbLf & "*********** TestHasKey tests ***********"
    Debug.Print True, hasKey(c, "a") ' True for scalar
    Debug.Print True, hasKey(c, "b") ' True for object
    Debug.Print True, hasKey(c, "A") ' True (even though case insensitive???)

    Debug.Print True, hasKey(Workbooks, ThisWorkbook.Name) ' True for non-collection type collections
    
    Dim d As New Dictionary
    d.Add "a", "a"
    d.Add "b", col("x", "y", "z")
    
    Debug.Print True, hasKey(d, "a") ' True for scalar
    Debug.Print True, hasKey(d, "b") ' True for object
    Debug.Print False, hasKey(d, "A") ' False - case sensitive by default
    
    Dim dObj As Object
    Set dObj = CreateObject("Scripting.Dictionary")
    
    dObj.Add "a", "a"
    dObj.Add "b", col("x", "y", "z")
    
    Debug.Print True, hasKey(dObj, "a") ' True for scalar
    Debug.Print True, hasKey(dObj, "b") ' True for object
    Debug.Print False, hasKey(dObj, "A") ' False - case sensitive by default
    
    ' Errors
    On Error Resume Next
        Err.Number = 0
        hasKey ThisWorkbook, "A"
        Debug.Print 9, Err.Number ' WorkBook doesn't have keys/items
        
        Err.Number = 0
        hasKey 5, "A"
        Debug.Print 9, Err.Number ' Variant doesn't have keys/items
    On Error GoTo 0

End Sub


Public Function hasKey(Container, key As Variant) As Boolean
    Dim ErrX As Integer
    Dim hasKeyFlag As Boolean
    Dim emptyFlag As Boolean
    
    ' First try .HasKey method on the object
    On Error Resume Next
        Err.Number = 0
        hasKeyFlag = Container.Exists(key)
        ErrX = Err.Number
    On Error GoTo 0
    If ErrX = 0 Then
        hasKey = hasKeyFlag
        Exit Function
    End If
    
    
    ' Then test with .Item method
    emptyFlag = False
    On Error Resume Next
        Err.Number = 0
        emptyFlag = TypeName(Container.Item(key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via .Item
        If emptyFlag Then ' Item was Empty/non-existant
            hasKey = False
        Else
            hasKey = True ' Item was not Empty
        End If
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 Then ' Retrieval Error, but .Item is correct access method stil. 424: Method not exist; 438: Compilation error
        hasKey = False
        Exit Function
    End If
    
    
    ' Then test with bracketed access, like ()
    emptyFlag = False
    On Error Resume Next
        Err.Number = 0
        emptyFlag = TypeName(Container(key)) = "Empty"
        ErrX = Err.Number
    On Error GoTo 0
    
    If ErrX = 0 Then ' No error trying to Access Key via ()
        If emptyFlag Then ' Item was Empty/non-existant
            hasKey = False
        Else
            hasKey = True ' Item was not Empty
        End If
        Exit Function
    ElseIf ErrX <> 424 And ErrX <> 438 And ErrX <> 13 Then ' Retrieval Error, but () is correct access method stil. 424: Method not exist; 438: Compilation error; 13: Variant bracketed ()
        hasKey = False
        Exit Function
    End If

    
    Dim errmsg As String
    On Error Resume Next
        errmsg = "Object"
        errmsg = errmsg & " of type '" & TypeName(Container) & "'"
        errmsg = errmsg & " have neither '.Exists' method, nor bracketed indexing '()', nor '.Item' method"
    On Error GoTo 0
    Err.Raise 9, , errmsg
    
End Function


'************"MiscNewKeys"


' this module is used to generate new keys to a container (collections, dict, sheets, etc)
' Use case is when we want to create a new sheet, but
' want to ensure we don't give a name that already exists in the workbook

Function NewSheetName(Name As String, Optional WB As Workbook)

    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' max 31 characters
    NewSheetName = Left(Name, 31)

    If Not Fn.hasKey(WB.Sheets, NewSheetName) Then
        ' sheet name doesn't exist, so we can continue
        Exit Function
    Else
        NewSheetName = GetNewKey(Name, WB.Sheets, 31)
    End If
End Function

Private Function TestGetNewKey()

    Dim c As New Collection, i As Long
    
    c.Add "bla", "name"
    For i = 1 To 100
        c.Add "bla", "name" & i
    Next i
    
    Debug.Print GetNewKey("name", c), "name101"
    Debug.Print GetNewKey("NewName", c), "NewName"

End Function


Function GetNewKey(Name As String, Container, Optional MaxLength As Long = -1, Optional depth As Long = 0) As String
    ' get a key that does not exists in a container (dict or collection)
    ' we keep appending, 1, 2, 3, ..., 10, 11 until the key is unique
    ' MaxLength is used when the key has a restriction on the maximum length
        ' for example sheet names can only be 31 characters long
    
    If MaxLength = -1 Then
        GetNewKey = Name
    Else
        GetNewKey = Left(Name, MaxLength)
    End If
    
    If Not hasKey(Container, GetNewKey) Then
        ' Key is "New" and we don't need further iteration
        Exit Function
    Else
        ' 31 max characters for sheet name
        depth = depth + 1
        If MaxLength = -1 Then
            GetNewKey = GetNewKey & depth
        Else
            GetNewKey = Left(GetNewKey, MaxLength - Len(CStr(depth))) & depth
        End If
        
        If Not hasKey(Container, GetNewKey) Then
            Exit Function
        End If
        
        GetNewKey = GetNewKey(Name, Container, MaxLength, depth)
    End If
End Function

'************"MiscRangeToArray"


' Converts a range to a normalized array.
Public Function RangeToArray(r As Range, _
                Optional IgnoreEmptyInFlatArray As Boolean) As Variant()
    ' vectors allocated to 1-dimensional arrays
    ' tables allocated to 2-dimensional array
    
    If r.Cells.Count = 1 Then
        RangeToArray = Array(r.Value)
    ElseIf r.Rows.Count = 1 Or r.Columns.Count = 1 Then
        RangeToArray = RangeTo1DArray(r, IgnoreEmptyInFlatArray)
    Else
        RangeToArray = r.Value
    End If
End Function



Function RangeTo1DArray( _
              r As Range _
            , Optional IgnoreEmpty As Boolean = True _
            ) As Variant()
    
    ' currently does the same as rangeToArray, just named better and is more efficient
    ' instead of reading from memory for every range item, we read it in only once
    
    Dim arr() As Variant ' the output array
    ReDim arr(r.Cells.Count - 1)
    
    Dim Values() As Variant ' values of the whole range
    If r.Cells.Count = 1 Then
        arr(0) = r.Value
        RangeTo1DArray = arr
        Exit Function
    End If
    
    Values = r.Value
    Dim i As Long, J As Long, counter As Long
    counter = 0
    For i = LBound(Values, 1) To UBound(Values, 1) ' rows
        For J = LBound(Values, 2) To UBound(Values, 2) ' columns
            If IsError(Values(i, J)) Then
                ' if error, we cannot check if empty, we need to add it
                arr(counter) = Values(i, J)
                counter = counter + 1
            ElseIf Values(i, J) = "" And IgnoreEmpty Then
                ReDim Preserve arr(UBound(arr) - 1) ' when there is an empty cell, just reduce array size by 1
            Else
                arr(counter) = Values(i, J)
                counter = counter + 1
            End If
        Next J
    Next i
    
    RangeTo1DArray = arr
    
End Function



'************"MiscRemoveGridLines"


Sub RemoveGridLines(WS As Worksheet)
    Dim view As WorksheetView
    For Each view In WS.Parent.Windows(1).SheetViews
        If view.Sheet.Name = WS.Name Then
            view.DisplayGridlines = False
            Exit Sub
        End If
    Next
End Sub

'************"MiscString"


Function randomString(length)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function


'************"MiscTables"


Function HasLO(Name As String, Optional WB As Workbook) As Boolean

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet, LO As ListObject
    
    For Each WS In WB.Sheets
        For Each LO In WS.ListObjects
            If Name = LO.Name Then
                HasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    HasLO = False

End Function


' get list object only using it's name from within a workbook
Function GetLO(Name As String, Optional WB As Workbook) As ListObject

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet, LO As ListObject
    
    For Each WS In WB.Sheets
        For Each LO In WS.ListObjects
            If Name = LO.Name Then
                Set GetLO = LO
                Exit Function
            End If
        Next LO
    Next WS
    
    If GetLO Is Nothing Then
        ' 9: Subscript out of range
        Err.Raise 9, , "List object '" & Name & "' not found in workbook '" & WB.Name & "'"
    End If

End Function

'************"MiscTableToDicts"


Private Sub TableToDictsTest()
    Dim Dicts As Collection
    Set Dicts = TableToDicts("TableToDictsTestData")
    ' read row 2 in column "b":
    Debug.Print Dicts(2)("b"), 5
End Sub

Function TableToDicts(TableName As String, Optional WB As Workbook) As Collection
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set TableToDicts = New Collection
    
    Dim d As Dictionary
    
    Dim Table As ListObject, lr As ListRow, lc As ListColumn
    Set Table = GetLO(TableName, WB)
    For Each lr In Table.ListRows
        Set d = New Dictionary
        For Each lc In Table.ListColumns
            d.Add lc.Name, lr.Range(1, lc.Index).Value
        Next lc
        
        TableToDicts.Add d
    Next lr
    
End Function

'************"Test__Dictionary"

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

'@TestMethod("Uncategorized")
Private Sub Test_dictget()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As Dictionary

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'************"Test__MiscAssign"

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

'@TestMethod("MiscAssign")
Private Sub Test_MiscAssign_variant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim i As Integer
    

    'Act:

    'Assert:
    Assert.AreEqual 5, assign(i, 5), "assign test succeeded"
    Assert.AreEqual 1.4, assign(i, 1.4), "assign test succeeded"
    
    
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscAssign")
Private Sub Test_MiscAssign_object()
    On Error GoTo TestFail
    
    'Arrange:
    Dim x As Variant
    Dim y As Variant
    Dim i As Variant
    Set i = col(4, 5, 6)
    assign x, i
    
    'Assert:
    Assert.AreEqual 4, x(1)
    Assert.AreEqual 5, assign(y, i)(2)


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'************"Test__MiscCollection"

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
    Dim c As Collection
    'Act:
    
    
    min c
    
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
    Dim c As Collection

    'Act:
    max c

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
    Dim c As Collection

    'Act:
    mean c

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

'************"Test__MiscCollectionCreate"

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

'@TestMethod("MiscCollectionCreate")
Private Sub Test_Col()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As Collection
    'Act:
    Set c = col(1, 3, 5)
    'Assert:
    'Assert.Succeed
    
    Assert.AreEqual 1, c(1), "col test succeeded"
    Assert.AreEqual 3, c(2), "col test succeeded"
    Assert.AreEqual 5, c(3), "col test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_zip()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim c1 As Collection
    Dim c2 As Collection
    Dim cOut As Collection

    'Act:
    Set c1 = col(1, 2, 3)
    Set c2 = col(4, 5, 6)
    
    Set cOut = zip(c1, c2)

    'Assert:
    Assert.AreEqual 1, cOut(1)(1), "zip test succeeded"
    Assert.AreEqual 4, cOut(1)(2), "zip test succeeded"
    
    Assert.AreEqual 2, cOut(2)(1), "zip test succeeded"
    Assert.AreEqual 5, cOut(2)(2), "zip test succeeded"
    
    Assert.AreEqual 3, cOut(3)(1), "zip test succeeded"
    Assert.AreEqual 6, cOut(3)(2), "zip test succeeded"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'************"Test__MiscHasKey"

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

'@TestMethod("MiscHasKey")
Private Sub test_HasKey_Collection()
    On Error GoTo TestFail
    
    
    'Arrange:
     Dim c As New Collection

    'Act:
    c.Add "foo", "a"
    c.Add col("x", "y", "z"), "b"
    
    'Assert:
    Assert.AreEqual True, hasKey(c, "a") ' True for scalar
    Assert.AreEqual True, hasKey(c, "b") ' True for scalar
    Assert.AreEqual True, hasKey(c, "A") ' True (even though case insensitive???)
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub test_HasKey_Workbook()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual True, hasKey(Workbooks, ThisWorkbook.Name)
    'Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub test_HasKey_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim d As New Dictionary
    
    'Act:
    d.Add "a", "foo"
    d.Add "b", col("x", "y", "z")

    'Assert:
    Assert.AreEqual True, hasKey(d, "a") ' True for scalar
    Assert.AreEqual True, hasKey(d, "b") ' True for scalar
    Assert.AreEqual False, hasKey(d, "A") ' False - case sensitive by default
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub test_HasKey_Dictionary_object()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dObj As Object
    Set dObj = CreateObject("Scripting.Dictionary")
    'Act:
    dObj.Add "a", "foo"
    dObj.Add "b", col("x", "y", "z")

    'Assert:
    Assert.AreEqual True, hasKey(dObj, "a") ' True for scalar
    Assert.AreEqual True, hasKey(dObj, "b") ' True for scalar
    Assert.AreEqual False, hasKey(dObj, "A") ' False - case sensitive by default

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MiscHasKey")
Private Sub test_HasKey_Dictionary_fail()                        'TODO Rename test
    Const ExpectedError As Long = 9              'TODO Change to expected error number
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    hasKey 5, "a"
    hasKey ThisWorkbook, "A"

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
