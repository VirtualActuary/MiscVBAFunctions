Attribute VB_Name = "MiscF"
Option Explicit

'************"aErrorEnums"
' Use a* so this is on top of the fn. MiscF library


Enum ErrNr
    '********************************************
    'These are internal error codes collected from
    'https://bettersolutions.com/vba/error-handling/error-codes.htm
    '********************************************
    ReturnWithoutGoSub = 3
    InvalidProcedureCall = 5
    Overflow = 6
    OutOfMemory_ = 7
    SubscriptOutOfRange = 9
    ThisArrayIsFixedOrTemporarilyLocked = 10
    DivisionByZero = 11
    TypeMismatch = 13
    OutOfStringSpace = 14
    ExpressionTooComplex = 16
    CantPerformRequestedOperation = 17
    UserInterruptOccurred = 18
    ResumeWithoutError = 20
    OutOfStackSpace = 28
    SubFunctionOrPropertyNotDefined = 35
    TooManyDLLApplicationClients = 47
    ErrorInLoadingDLL = 48
    BadDLLCallingConvention = 49
    InternalError = 51
    BadFileNameOrNumber = 52
    FileNotFound = 53
    BadFileMode = 54
    FileAlreadyOpen = 55
    DeviceIOError = 57
    FileAlreadyExists = 58
    BadRecordLength = 59
    DiskFull = 61
    InputPastEndOfFile = 62
    BadRecordNumber = 63
    TooManyFiles = 67
    DeviceUnavailable = 68
    PermissionDenied = 70
    DiskNotReady = 71
    CantRenameWithDifferentDrive = 74
    PathFileAccessError = 75
    PathNotFound = 76
    ObjectVariableOrWithBlockVariableNotSet = 91
    ForLoopNotInitialized = 92
    InvalidPatternString = 93
    InvalidUseOfNull = 94
    CantCallFriendProcedureOnAnObjectThatIsNotAnInstanceOfTheDefiningClass = 97
    SystemDLLCouldNotBeLoaded = 298
    CantUseCharacterDeviceNamesInSpecifiedFileNames = 320
    InvalidFileFormat = 321
    CantCreateNecessaryTemporaryFile = 322
    InvalidFormatInResourceFile = 325
    DataValueNamedNotFound = 327
    IllegalParameterCantWriteArrays = 328
    CouldNotAccessSystemRegistry = 335
    ActiveXComponentNotCorrectlyRegistered = 336
    ActiveXComponentNotFound = 337
    ActiveXComponentDidNotRunCorrectly = 338
    ObjectAlreadyLoaded = 360
    CantLoadOrUnloadThisObject = 361
    ActiveXControlSpecifiedNotFound = 363
    ObjectWasUnloaded = 364
    UnableToUnloadWithinThisContext = 365
    TheSpecifiedFileIsOutOfDateThisProgramRequiresALaterVersion = 368
    TheSpecifiedObjectCantBeUsedAsAnOwnerFormForShow = 371
    InvalidPropertyValue = 380
    InvalidPropertyarrayIndex = 381
    PropertySetCantBeExecutedAtRunTime = 382
    PropertySetCantBeUsedWithAReadonlyProperty = 383
    NeedPropertyArrayIndex = 385
    PropertySetNotPermitted = 387
    PropertyGetCantBeExecutedAtRunTime = 393
    PropertyGetCantBeExecutedOnWriteonlyProperty = 394
    FormAlreadyDisplayedCantShowModally = 400
    CodeMustCloseTopmostModalFormFirst = 402
    PermissionToUseObjectDenied = 419
    PropertyNotFound = 422
    PropertyOrMethodNotFound = 423
    ObjectRequired = 424
    InvalidObjectUse = 425
    ActiveXComponentCantCreateObjectOrReturnReferenceToThisObject = 429
    ClassDoesntSupportAutomation = 430
    FileNameOrClassNameNotFoundDuringAutomationOperation = 432
    ObjectDoesntSupportThisPropertyOrMethod = 438
    AutomationError = 440
    ConnectionToTypeLibraryOrObjectLibraryForRemoteProcessHasBeenLost = 442
    AutomationObjectDoesntHaveADefaultValue = 443
    ObjectDoesntSupportThisAction = 445
    ObjectDoesntSupportNamedArguments = 446
    ObjectDoesntSupportCurrentLocaleSetting = 447
    NamedArgumentNotFound = 448
    ArgumentNotOptionalOrInvalidPropertyAssignment = 449
    WrongNumberOfArgumentsOrInvalidPropertyAssignment = 450
    ObjectNotACollection = 451
    InvalidOrdinal = 452
    SpecifiedDLLFunctionNotFound = 453
    CodeResourceNotFound = 454
    CodeResourceLockError = 455
    ThisKeyIsAlreadyAssociatedWithAnElementOfThisCollection = 457
    VariableUsesATypeNotSupportedInVisualBasic = 458
    ThisComponentDoesntSupportEvents = 459
    InvalidClipboardFormat = 460
    SpecifiedFormatDoesntMatchFormatOfData = 461
    CantCreateAutoRedrawImage = 480
    InvalidPicture = 481
    PrinterError = 482
    PrinterDriverDoesNotSupportSpecifiedProperty = 483
    ProblemGettingPrinterInformationFromTheSystemMakeSureThePrinterIsSetUpCorrectly = 484
    InvalidPictureType = 485
    CantPrintFormImageToThisTypeOfPrinter = 486
    CantEmptyClipboard = 520
    CantOpenClipboard = 521
    CantSaveFileToTEMPDirectory = 735
    SearchTextNotFound = 744
    ReplacementsTooLong = 746
    OutOfMemory = 31001
    NoObject = 31004
    ClassIsNotSet = 31018
    UnableToActivateObject = 31027
    UnableToCreateEmbeddedObject = 31032
    ErrorSavingToFile = 31036
End Enum

'************"aFSO"
' allows us to use FSO functions anywhere in the project
' Use a* so this is on top of the fn. MiscF library

Public fso As New FileSystemObject

'************"Casing"
' Uncomment and comment block to get casing back for the project
'
'
'Dim J
'Dim I
'Dim M
'Dim WB
'Dim WS
'Dim dict ' for now... :/
'Dim r ' for now... :/
'Dim Keys
'Dim Columns
'Dim Values
'Dim Column
'Dim ColumnName

'************"MiscArray"
'@IgnoreModule ImplicitByRefModifier


' Functions for 1D and 2D arrays only.
' Replaces all Errors in the input array with vbNullString.
' The input array is modified (pass by referance) and the function returns the array
Public Function ErrorToNullStringTransformation(tableArr() As Variant) As Variant
    If is2D(tableArr) Then
        ErrorToNullStringTransformation = ErrorToNull2D(tableArr)
    Else
        ErrorToNullStringTransformation = ErrorToNull1D(tableArr)
    End If
End Function


' Functions for 1D and 2D arrays only.
' Converts the decimal seperator in the float input to a "." for each entry in the input array
' and returns the result as a string.
' Only works when converting from the system's decimal seperator.
' Custom seperators not supported.
' The input array is modified (pass by referance) and the function returns the array.
Public Function EnsureDotSeparatorTransformation(tableArr() As Variant) As Variant
    If is2D(tableArr) Then
        EnsureDotSeparatorTransformation = EnsureDotSeparator2D(tableArr)
    Else
        EnsureDotSeparatorTransformation = EnsureDotSeparator1D(tableArr)
    End If
End Function


' Functions for 1D and 2D arrays only.
' Converts all Date/DateTime entries in the input array to string.
' The input array is modified (pass by referance) and the function returns the array.
Public Function DateToStringTransformation(tableArr() As Variant, Optional fmt As String = "yyyy-mm-dd") As Variant
    If is2D(tableArr) Then
        DateToStringTransformation = DateToString2D(tableArr, fmt)
    Else
        DateToStringTransformation = DateToString1D(tableArr, fmt)
    End If
End Function


' Check if a collection is 1D or 2D.
' 3D is not supported
Private Function is2D(arr As Variant)
    On Error GoTo Err
    is2D = (UBound(arr, 2) - LBound(arr, 2) > 1)
    Exit Function
Err:
    is2D = False
End Function


Private Function dateToString(d As Date, fmt As String) As String
    dateToString = Format(d, fmt)
End Function


Private Function decStr(x As Variant) As String
     decStr = CStr(x)

     'Frikin ridiculous loops for VBA
     If IsNumeric(x) Then
        decStr = Replace(decStr, Format(0, "."), ".")
        ' Format(0, ".") gives the system decimal separator
     End If

End Function


Private Function ErrorToNull2D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsError(tableArr(I, J)) Then ' set all error values to an empty string
                tableArr(I, J) = vbNullString
            End If
        Next J
    Next I
    ErrorToNull2D = tableArr
End Function


Private Function ErrorToNull1D(tableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsError(tableArr(I)) Then ' set all error values to an empty string
            tableArr(I) = vbNullString
        End If
    Next I
    ErrorToNull1D = tableArr
End Function


Private Function EnsureDotSeparator2D(tableArr As Variant) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsNumeric(tableArr(I, J)) Then ' force numeric values to use . as decimal separator
                tableArr(I, J) = decStr(tableArr(I, J))
            End If
        Next J
    Next I
    EnsureDotSeparator2D = tableArr
End Function


Private Function EnsureDotSeparator1D(tableArr As Variant) As Variant
    Dim I As Long
    For I = LBound(tableArr) To UBound(tableArr)
        If IsNumeric(tableArr(I)) Then ' force numeric values to use . as decimal separator
            tableArr(I) = decStr(tableArr(I))
        End If
    Next I
    EnsureDotSeparator1D = tableArr
End Function


Private Function DateToString2D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long, J As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        For J = LBound(tableArr, 2) To UBound(tableArr, 2)
            If IsDate(tableArr(I, J)) Then ' format dates as strings to avoid some user's stupid default date settings
                tableArr(I, J) = dateToString(CDate(tableArr(I, J)), fmt)
            End If
        Next J
    Next I
    DateToString2D = tableArr
End Function


Private Function DateToString1D(tableArr As Variant, fmt As String) As Variant
    Dim I As Long
    For I = LBound(tableArr, 1) To UBound(tableArr, 1)
        If IsDate(tableArr(I)) Then ' format dates as strings to avoid some user's stupid default date settings
            tableArr(I) = dateToString(CDate(tableArr(I)), fmt)
        End If
    Next I
    DateToString1D = tableArr
End Function

'************"MiscAssign"
' Assign a value to a variable and also return that value. The goal of this function is to
' overcome the different `set` syntax for assigning an object vs. assigning a native type
' like Int, Double etc. Additionally this function has similar functionality to Python's
' walrus operator: https://towardsdatascience.com/the-walrus-operator-7971cd339d7d


Public Function assign(ByRef var As Variant, ByRef val As Variant)
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


Function IsValueInCollection(col As Collection, val As Variant, Optional CaseSensitive As Boolean = False) As Boolean
    Dim ValI As Variant
    For Each ValI In col
        ' only check if not an object:
        If Not IsObject(ValI) Then
            If CaseSensitive Then
                IsValueInCollection = ValI = val
            Else
                IsValueInCollection = LCase(ValI) = LCase(val)
            End If
            ' exit if found
            If IsValueInCollection Then Exit Function
        End If
    Next ValI
End Function



'************"MiscCollectionCreate"


Public Function col(ParamArray Args() As Variant) As Collection
    Set col = New Collection
    Dim I As Long

    For I = LBound(Args) To UBound(Args)
        col.Add Args(I)
    Next

End Function


Public Function zip(ParamArray Args() As Variant) As Collection
    Dim I As Long
    Dim J As Long
    Dim M As Long
    
    M = -1
    For I = LBound(Args) To UBound(Args)
        If M = -1 Then
            M = Args(I).Count
        ElseIf Args(I).Count < M Then
            M = Args(I).Count
        End If
    Next I

    Set zip = New Collection
    Dim ICol As Collection
    For I = 1 To M
        Set ICol = New Collection
        For J = LBound(Args) To UBound(Args)
            ICol.Add Args(J).Item(I)
        Next J
        zip.Add ICol
    Next I
End Function




'************"MiscCollectionSort"



Private Sub TestBubbleSort()
    Dim coll As Collection
    Set coll = col("variables10", "variables", "variables2", "variables_10", "variables_2")
    Set coll = BubbleSort(coll)
    
    Debug.Print coll(1), "variables"
    Debug.Print coll(2), "variables10" ' :/
    Debug.Print coll(3), "variables2" ' :/
    Debug.Print coll(4), "variables_10" ' :/
    Debug.Print coll(5), "variables_2" ' :/
    
End Sub


Public Function BubbleSort(coll As Collection) As Collection
    
    ' from: https://github.com/austinleedavis/VBA-utilities/blob/f23f1096d8df0dfdc740e5a3bec36525d61a3ffc/Collections.bas#L73
    ' this is an easy implementation but a slow sorting algorithm
    ' do not use for large collections
    
    Dim SortedColl As Collection
    Set SortedColl = New Collection
    Dim vItm As Variant
    ' copy the collection"
    For Each vItm In coll
        SortedColl.Add vItm
    Next vItm

    Dim I As Long, J As Long
    Dim vTemp As Variant

    'Two loops to bubble sort
    For I = 1 To SortedColl.Count - 1
        For J = I + 1 To SortedColl.Count
            If SortedColl(I) > SortedColl(J) Then ' 1 = I is larger than J
                'store the lesser item
               assign vTemp, SortedColl(J) ' assign
                'remove the lesser item
               SortedColl.Remove J
                're-add the lesser item before the greater Item
               SortedColl.Add vTemp, , I
            End If
        Next J
    Next I
    
    Set BubbleSort = SortedColl
    
End Function


'************"MiscCreateTextFile"


Private Sub testCreateTextFile()
    CreateTextFile "foo", ThisWorkbook.Path & "\tests\MiscCreateTextFile\test.txt"
    ' TODO: assertion
End Sub

Public Sub CreateTextFile(ByVal Content As String, ByVal FilePath As String)
    ' Creates a new / overwrites an existing text file with Content
    
    Dim oFile As Integer
    oFile = FreeFile
    
    Open FilePath For Output As #oFile
        Print #oFile, Content
    Close #oFile

End Sub

'************"MiscDictionary"
' This module contains functions to retrieve values given a
' key in a dictionary



'@IgnoreModule ImplicitByRefModifier

Private Sub testDictget()

    Dim d As Dictionary
    Set d = dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(d, "a"), 2 ' returns 2
    Debug.Print dictget(d, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(d, "c", vbNullString), vbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(d, "c")
        Debug.Print Err.Number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Sub


Public Function dictget(d As Dictionary, key As Variant, Optional default As Variant = Empty) As Variant
'   Params:
'       d: Dictionary to read the value from...
'       key: The key ....
'
'   Returns:
'
    
    
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
        
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, errmsg)
    End If
End Function

'************"MiscDictionaryCreate"


Public Function dict(ParamArray Args() As Variant) As Dictionary
    'Case sensitive dictionary
    
    Dim errmsg As String
    Set dict = New Dictionary
    
    Dim I As Long
    Dim Cnt As Long
    Cnt = 0
    For I = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If I + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(I) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dict.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function


Public Function dicti(ParamArray Args() As Variant) As Dictionary
    'Case insensitive dictionary
    
    Dim errmsg As String
    Set dicti = New Dictionary
    dicti.CompareMode = TextCompare
    
    Dim I As Long
    Dim Cnt As Long
    Cnt = 0
    For I = LBound(Args) To UBound(Args)
        Cnt = Cnt + 1
        If (Cnt Mod 2) = 0 Then GoTo Cont

        If I + 1 > UBound(Args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & Args(I) & "`": On Error GoTo 0
            Err.Raise 9, , errmsg
        End If
        
        dicti.Add Args(I), Args(I + 1)
Cont:
    Next I

End Function


'************"MiscEarlyBindings"
'@IgnoreModule ImplicitByRefModifier



' Add references for this project programatically. If you are uncertain what to put here,
' Go to Tools -> References and use the filename of the reference (eg. msado15.dll for
' Microsoft ActiveX Data Objects 6.1 Library'), then run getPackageGUID("msado15.dll")
' to see what options you have:
'**********************************************************************************
'* Add selected references to this project
'**********************************************************************************
Sub addEarlyBindings()
    On Error GoTo ErrorHandler
        If Not isBindingNameLoaded("ADODB") Then
            'Microsoft ActiveX Data Objects 6.1
            ThisWorkbook.VBProject.References.addFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6.1, 0
        End If
        
        If Not isBindingNameLoaded("VBIDE") Then
            'Microsoft Visual Basic for Applications Extensibility 5.3
            ThisWorkbook.VBProject.References.addFromGuid "{0002E157-0000-0000-C000-000000000046}", 5.3, 0
        End If
        
        
        If Not isBindingNameLoaded("Scripting") Then
            'Microsoft Scripting Runtime version 1.0
            ThisWorkbook.VBProject.References.addFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        End If
        
    
        If Not isBindingNameLoaded("VBScript_RegExp_55") Then
            'Microsoft VBScript Regular Expressions 5.5
            ThisWorkbook.VBProject.References.addFromGuid "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5
        End If
        
        If Not isBindingNameLoaded("Shell32") Then
            'Microsoft Shell Controls And Automation
            ThisWorkbook.VBProject.References.addFromGuid "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}", 1, 0
        End If
                
    Exit Sub
ErrorHandler:
End Sub


'**********************************************************************************
'* Verify if a reference is loaded
'**********************************************************************************
Private Function isBindingNameLoaded(ref As String) As Boolean
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    isBindingNameLoaded = False
    Dim xRef As Variant
    For Each xRef In ThisWorkbook.VBProject.References
        If LCase(xRef.Name) = LCase(ref) Then
            isBindingNameLoaded = True
        End If
    Next xRef
    
End Function


'**********************************************************************************
'* Print all current active GUIDs
'**********************************************************************************
Private Sub printAllEarlyBindings()
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    Dim xRef As Variant
    For Each xRef In ThisWorkbook.VBProject.References
        Debug.Print "**************" & xRef.Name
        Debug.Print xRef.Description
        Debug.Print xRef.Major
        Debug.Print xRef.Minor
        Debug.Print xRef.FullPath
        Debug.Print xRef.GUID
        Debug.Print ""
    Next xRef
    
End Sub

'************"MiscEnsureDictIUtil"
'@IgnoreModule ImplicitByRefModifier


Function EnsureDictI(Container As Variant) As Object
    Dim key As Variant
    Dim Item As Variant
    
    If TypeOf Container Is Collection Then
        Dim c As Collection
        Set c = New Collection
        
        For Each Item In Container
            If TypeOf Item Is Collection Or TypeOf Item Is Dictionary Then
                c.Add EnsureDictI(Item)
            Else
                c.Add Item
            End If
        Next Item
        
        Set EnsureDictI = c
        
    ElseIf TypeOf Container Is Dictionary Then
        Dim d As Dictionary
        Set d = New Dictionary
        d.CompareMode = TextCompare
        
        For Each key In Container.Keys
            If TypeOf Container.Item(key) Is Collection Or TypeOf Container.Item(key) Is Dictionary Then
                d.Add key, EnsureDictI(Container.Item(key))
            Else
                d.Add key, Container.Item(key)
            End If
        Next key
        
        Set EnsureDictI = d
    Else
    
        Dim errmsg As String
        errmsg = "ConvertToTextCompare only supports type 'Dictionary' and 'Collection'"
        On Error Resume Next: errmsg = errmsg & ". Got type '" & TypeName(Container) & "'": On Error GoTo 0
        Err.Raise 5, , errmsg
        
    End If
End Function

'************"MiscErrorMessage"

' #####################################################################
' ##### This module is version controlled in RunnerModule/Modules #####
' #####################################################################

'@Folder("error handling")

Function ErrorMessage(ErrorCode As Integer, _
                              Optional SubMessage As String) As String

    Dim M As String
    Select Case ErrorCode
        '**********************************************
        'VBA messages from https://bettersolutions.com/vba/error-handling/error-codes.htm
        '**********************************************
        Case 3: M = "Return without GoSub"
        Case 5: M = "Invalid procedure call"
        Case 6: M = "Overflow"
        Case 7: M = "Out of memory"
        Case 9: M = "Subscript out of range"
        Case 10: M = "This array is fixed or temporarily locked"
        Case 11: M = "Division by zero"
        Case 13: M = "Type mismatch"
        Case 14: M = "Out of string space"
        Case 16: M = "Expression too complex"
        Case 17: M = "Can't perform requested operation"
        Case 18: M = "User interrupt occurred"
        Case 20: M = "Resume without error"
        Case 28: M = "Out of stack space"
        Case 35: M = "Sub, Function, or Property not defined"
        Case 47: M = "Too many DLL application clients"
        Case 48: M = "Error in loading DLL"
        Case 49: M = "Bad DLL calling convention"
        Case 51: M = "Internal error"
        Case 52: M = "Bad file name or number"
        Case 53: M = "File not found"
        Case 54: M = "Bad file mode"
        Case 55: M = "File already open"
        Case 57: M = "Device I/O error"
        Case 58: M = "File already exists"
        Case 59: M = "Bad record length"
        Case 61: M = "Disk full"
        Case 62: M = "Input past end of file"
        Case 63: M = "Bad record number"
        Case 67: M = "Too many files"
        Case 68: M = "Device unavailable"
        Case 70: M = "Permission denied"
        Case 71: M = "Disk not ready"
        Case 74: M = "Can't rename with different drive"
        Case 75: M = "Path/File access error"
        Case 76: M = "Path not found"
        Case 91: M = "Object variable or With block variable not set"
        Case 92: M = "For loop not initialized"
        Case 93: M = "Invalid pattern string"
        Case 94: M = "Invalid use of Null"
        Case 97: M = "Can't call Friend procedure on an object that is not an instance of the defining class"
        Case 298: M = "System DLL could not be loaded"
        Case 320: M = "Can't use character device names in specified file names"
        Case 321: M = "Invalid file format"
        Case 322: M = "Can't create necessary temporary file"
        Case 325: M = "Invalid format in resource file"
        Case 327: M = "Data value named not found"
        Case 328: M = "Illegal parameter; can't write arrays"
        Case 335: M = "Could not access system registry"
        Case 336: M = "ActiveX component not correctly registered"
        Case 337: M = "ActiveX component not found"
        Case 338: M = "ActiveX component did not run correctly"
        Case 360: M = "Object already loaded"
        Case 361: M = "Can't load or unload this object"
        Case 363: M = "ActiveX control specified not found"
        Case 364: M = "Object was unloaded"
        Case 365: M = "Unable to unload within this context"
        Case 368: M = "The specified file is out of date. This program requires a later version"
        Case 371: M = "The specified object can't be used as an owner form for Show"
        Case 380: M = "Invalid property value"
        Case 381: M = "Invalid property-array index"
        Case 382: M = "Property Set can't be executed at run time"
        Case 383: M = "Property Set can't be used with a read-only property"
        Case 385: M = "Need property-array index"
        Case 387: M = "Property Set not permitted"
        Case 393: M = "Property Get can't be executed at run time"
        Case 394: M = "Property Get can't be executed on write-only property"
        Case 400: M = "Form already displayed; can't show modally"
        Case 402: M = "Code must close topmost modal form first"
        Case 419: M = "Permission to use object denied"
        Case 422: M = "Property not found"
        Case 423: M = "Property or method not found"
        Case 424: M = "Object required"
        Case 425: M = "Invalid object use"
        Case 429: M = "ActiveX component can't create object or return reference to this object"
        Case 430: M = "Class doesn't support Automation"
        Case 432: M = "File name or class name not found during Automation operation"
        Case 438: M = "Object doesn't support this property or method"
        Case 440: M = "Automation error"
        Case 442: M = "Connection to type library or object library for remote process has been lost"
        Case 443: M = "Automation object doesn't have a default value"
        Case 445: M = "Object doesn't support this action"
        Case 446: M = "Object doesn't support named arguments"
        Case 447: M = "Object doesn't support current locale setting"
        Case 448: M = "Named argument not found"
        Case 449: M = "Argument not optional or invalid property assignment"
        Case 450: M = "Wrong number of arguments or invalid property assignment"
        Case 451: M = "Object not a collection"
        Case 452: M = "Invalid ordinal"
        Case 453: M = "Specified DLL function not found"
        Case 454: M = "Code resource not found"
        Case 455: M = "Code resource lock error"
        Case 457: M = "This key is already associated with an element of this collection"
        Case 458: M = "Variable uses a type not supported in Visual Basic"
        Case 459: M = "This component doesn't support events"
        Case 460: M = "Invalid Clipboard format"
        Case 461: M = "Specified format doesn't match format of data"
        Case 480: M = "Can't create AutoRedraw image"
        Case 481: M = "Invalid picture"
        Case 482: M = "Printer error"
        Case 483: M = "Printer driver does not support specified property"
        Case 484: M = "Problem getting printer information from the system. Make sure the printer is set up correctly"
        Case 485: M = "Invalid picture type"
        Case 486: M = "Can't print form image to this type of printer"
        Case 520: M = "Can't empty Clipboard"
        Case 521: M = "Can't open Clipboard"
        Case 735: M = "Can't save file to TEMP directory"
        Case 744: M = "Search text not found"
        Case 746: M = "Replacements too long"
        Case 31001: M = "Out of memory"
        Case 31004: M = "No object"
        Case 31018: M = "Class is not set"
        Case 31027: M = "Unable to activate object"
        Case 31032: M = "Unable to create embedded object"
        Case 31036: M = "Error saving to file"
        Case 31037: M = "Error loading from file"
        Case Else: M = "Unknown error"
    End Select
    
    ErrorMessage = M & IIf(SubMessage = "", "", ": " & SubMessage)
    
    ErrorMessage = BreakLines(ErrorMessage)
    
End Function

Function BreakLines(SubMessage As String) As String
    
    Dim MsgArr() As String, I As Integer
    ' split by spaces
    MsgArr = Split(SubMessage, " ")
    
    For I = LBound(MsgArr) To UBound(MsgArr)
        ' only split if one "word" is longer than 100 or 50 characters
        If Len(MsgArr(I)) > 100 Then
            MsgArr(I) = BreakLines & Mid(MsgArr(I), 1, 50) & chr(10) & Mid(MsgArr(I), 51, 50) & chr(10) & right(MsgArr(I), Len(MsgArr(I)) - 50 * 2)
        ElseIf Len(MsgArr(I)) > 50 Then
            MsgArr(I) = Mid(MsgArr(I), 1, 50) & chr(10) & right(MsgArr(I), Len(MsgArr(I)) - 50)
        End If
    Next I
    
    ' join back with spaces
    BreakLines = Join(MsgArr, " ")
    
End Function

'************"MiscExcel"




Public Function ExcelBook( _
      Path As String _
    , Optional MustExist As Boolean = False _
    , Optional ReadOnly As Boolean = False _
    , Optional SaveOnError As Boolean = False _
    , Optional CloseOnError As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    
    On Error GoTo finally
    
    If fso.FileExists(Path) Then
    
        Set ExcelBook = OpenWorkbook(Path, ReadOnly)
    
    Else
        
        If MustExist Then
            Err.Raise -999, , "FileNotFoundError: File '" & fso.GetAbsolutePathName(Path) & "' does not exist."
        Else
            Set ExcelBook = Workbooks.Add
            
            If SaveOnError Then
                ExcelBook.SaveAs Path
            End If
        End If
        
    End If
    
    Exit Function
    
finally:
    If SaveOnError Then
        ExcelBook.Save
    End If
    
    If CloseOnError Then
        ExcelBook.Close (False)
    End If
    
End Function

Function OpenWorkbook( _
      Path As String _
    , Optional ReadOnly As Boolean = False _
    ) As Workbook
    
    
    If hasKey(Workbooks, fso.GetFileName(Path)) Then
        Set OpenWorkbook = Workbooks(fso.GetFileName(Path))
        
        ' check if the workbook is actually the one specified in path
        ' use AbsolutePathName to remove any relative path references  (\..\ / \.\)
        If LCase(OpenWorkbook.FullName) <> LCase(fso.GetAbsolutePathName(Path)) Then
            Debug.Print fso.GetAbsolutePathName(Path)
            Err.Raise 457, , "Existing workbook with the same name is already open: '" & fso.GetFileName(Path) & "'"
        End If
        
        If ReadOnly And OpenWorkbook.ReadOnly = False Then
            Err.Raise -999, , "Workbook'" & fso.GetFileName(Path) & "' is already open and is not in ReadOnly mode. Only closed workbooks can be opened as readonly."
        End If
    Else
        Set OpenWorkbook = Workbooks.Open(Path, ReadOnly:=ReadOnly)
    End If
End Function


'************"MiscFreezePanes"
'@IgnoreModule ImplicitByRefModifier



Private Sub test()
    
    
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Worksheets(1)
    FreezePanes WS.Range("D6")
    
    
End Sub

Public Sub FreezePanes(r As Range)
    
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

Public Sub UnFreezePanes(WS As Worksheet)
    
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
'@IgnoreModule ImplicitByRefModifier


Private Sub TestGetUniqueItems()
    Dim arr(3) As Variant
    
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
    
End Sub

Public Function GetUniqueItems(arr() As Variant, _
            Optional CaseSensitive As Boolean = True) As Variant
    If ArrayLen(arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim d As New Dictionary
        If Not CaseSensitive Then
            d.CompareMode = TextCompare
        End If
        
        Dim I As Long
        For I = LBound(arr) To UBound(arr)
            If Not d.Exists(arr(I)) Then
                d.Add arr(I), arr(I)
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
'@IgnoreModule ImplicitByRefModifier


Public Sub GroupRowsOnIndentations(r As Range)
    ' groups the rows based on indentations of the cells in the range

    Dim ri As Range
    For Each ri In r
        ri.EntireRow.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub

Public Sub GroupColumnsOnIndentations(r As Range)
    ' groups the columns based on indentations of the cells in the range

    Dim ri As Range
    For Each ri In r
        ri.EntireColumn.OutlineLevel = ri.IndentLevel + 1
    Next ri
    
End Sub


Private Sub TestRemoveGroupings()
    ' Test rows
    RemoveRowGroupings ThisWorkbook.Worksheets("GroupOnIndentations")
    ' Test columns
    RemoveColumnGroupings ThisWorkbook.Worksheets("GroupOnIndentations")
End Sub


Public Sub RemoveRowGroupings(WS As Worksheet)
    Dim r As Range
    Dim ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels RowLevels:=8
    For Each ri In r.Columns(1)
        ri.EntireRow.OutlineLevel = 1
    Next ri
End Sub

Public Sub RemoveColumnGroupings(WS As Worksheet)
    Dim r As Range
    Dim ri As Range
    Set r = WS.UsedRange ' todo: better way to find last "active" cell
    WS.Outline.ShowLevels columnlevels:=8
    For Each ri In r.Rows(1)
        ri.EntireColumn.OutlineLevel = 1
    Next ri
End Sub


'************"MiscHasKey"
' Functions to check whether a key exists in a container


'@IgnoreModule ImplicitByRefModifier

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


Public Function hasKey(Container As Variant, key As Variant) As Boolean
    ' Checks whether a key exists in an existing container
    ' The container can be a `Collection`, `Dictionary` or any
    ' built-in Dictionary-like object. For example `ThisWorkbook.Sheets`
    '
    ' Args:
    '     Container: The container in which to look for the key
    '     key: The key to look for in the container
    '
    ' Returns:
    '     True for success, False otherwise
    
    
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
        hasKey = Not emptyFlag
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
        hasKey = Not emptyFlag
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
'@IgnoreModule ImplicitByRefModifier


' this module is used to generate new keys to a container (collections, dict, sheets, etc)
' Use case is when we want to create a new sheet, but
' want to ensure we don't give a name that already exists in the workbook

Public Function NewSheetName(Name As String, Optional WB As Workbook)

    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' max 31 characters
    NewSheetName = Left(Name, 31)

    If Not hasKey(WB.Sheets, NewSheetName) Then
        ' sheet name doesn't exist, so we can continue
        Exit Function
    Else
        NewSheetName = GetNewKey(Name, WB.Sheets, 31)
    End If
End Function

Private Sub TestGetNewKey()

    Dim c As New Collection
    Dim I As Long
    
    c.Add "bla", "name"
    For I = 1 To 100
        c.Add "bla", "name" & I
    Next I
    
    Debug.Print GetNewKey("name", c), "name101"
    Debug.Print GetNewKey("NewName", c), "NewName"

End Sub


Public Function GetNewKey(Name As String, Container As Variant, Optional MaxLength As Long = -1, Optional depth As Long = 0) As String
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

'************"MiscPowerQuery"
'@IgnoreModule ImplicitByRefModifier


' Helpful functions to help with Power Query manipulations in VBA

Private Sub MiscPowerQueryTests()
    Debug.Print doesQueryExist("foo"), False
End Sub


Public Function doesQueryExist(ByVal queryName As String, Optional WB As Workbook) As Boolean
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Helper function to check if a query with the given name already exists
    Dim qry As WorkbookQuery
    For Each qry In WB.queries
        If (qry.Name = queryName) Then
            doesQueryExist = True
            Exit Function
        End If
    Next
    doesQueryExist = False
End Function


Public Function getQuery(Name As String, Optional WB As Workbook) As WorkbookQuery
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim qry As WorkbookQuery
    For Each qry In WB.queries
        If qry.Name = Name Then
            Set getQuery = qry
            Exit Function
        End If
    Next qry
    
    Err.Raise 999, , "Query " & Name & " does not exist"
    
End Function


Public Function updateQuery(Name As String, queryFormula As String, Optional WB As Workbook) As WorkbookQuery
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a query to the new formula
    ' if the query doesn't exist, a new one is created
    
    If doesQueryExist(Name, WB) Then
        Set updateQuery = getQuery(Name, WB)
        updateQuery.formula = queryFormula
    Else
        Set updateQuery = WB.queries.Add(Name, queryFormula)
    End If
    
End Function

Public Function updateQueryAndRefreshListObject(Name As String, queryFormula As String, Optional WB As Workbook) As WorkbookQuery
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a power query query
    ' Also waits for the query to refresh before continuing the code
    
    ' assumes the ListObject and Query has the same name
    Set updateQueryAndRefreshListObject = updateQuery(Name, queryFormula, WB)
    
    WaitForListObjectRefresh Name, WB
    
End Function


Public Sub WaitForListObjectRefresh(Name As String, Optional WB As Workbook)
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Refreshes the query before continuing the code
    
    Dim LO As ListObject
    Set LO = GetLO(Name, WB)
    Dim BGRefresh As Boolean
    With LO.QueryTable
        BGRefresh = .BackgroundQuery
        .BackgroundQuery = False
        .Refresh
        .BackgroundQuery = BGRefresh
    End With
    
End Sub

Public Sub loadToWorkbook(queryName As String, Optional WB As Workbook)
    
    ' loads a query to a sheet in the workbook
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim LO As ListObject
    If HasLO(queryName, WB) Then
        Set LO = GetLO(queryName, WB)
        LO.Refresh
    Else
        Dim WS As Worksheet
        Set WS = WB.Worksheets.Add(After:=ActiveSheet)
        WS.Name = NewSheetName(queryName, ThisWorkbook)
        
        With WS.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & queryName & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .ListObject.DisplayName = queryName
            .Refresh BackgroundQuery:=False
        End With
        
    End If
    
End Sub

Function addToWorkbookConnections(Query As WorkbookQuery, Optional WB As Workbook) As WorkbookConnection
    ' adds a query to workbookconnections so that it can be used in pivot tables
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim ConnectionName As String, CommandString As String, CommandText As String, CommandType
    ConnectionName = "Query - " & Query.Name
    CommandString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Query.Name & ";Extended Properties="""""
    CommandText = "SELECT * FROM [" & Query.Name & "]"
    CommandType = 2
    
    ' This code loads the query to the workbook connections
    If hasKey(WB.Connections, ConnectionName) Then
        Set addToWorkbookConnections = WB.Connections(ConnectionName)
        addToWorkbookConnections.OLEDBConnection.Connection = CommandString
        addToWorkbookConnections.OLEDBConnection.CommandText = CommandText
        addToWorkbookConnections.OLEDBConnection.CommandType = CommandType
    Else
        Set addToWorkbookConnections = _
        WB.Connections.Add2(ConnectionName, _
            "Connection to the '" & Query.Name & "' query in the workbook.", _
            CommandString _
            , CommandText, CommandType)
        ' should not be loaded to the data model, else we cannot link two pivots to the same cache linking from this query
    End If

End Function



Sub refreshAllQueriesAndPivots(Optional WB As Workbook)
    If WB Is Nothing Then Set WB = ThisWorkbook
    WB.RefreshAll
End Sub



'************"MiscRangeToArray"
'@IgnoreModule ImplicitByRefModifier


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



Public Function RangeTo1DArray( _
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
    Dim I As Long
    Dim J As Long
    Dim counter As Long
    counter = 0
    For I = LBound(Values, 1) To UBound(Values, 1) ' rows
        For J = LBound(Values, 2) To UBound(Values, 2) ' columns
            If IsError(Values(I, J)) Then
                ' if error, we cannot check if empty, we need to add it
                arr(counter) = Values(I, J)
                counter = counter + 1
            ElseIf Values(I, J) = vbNullString And IgnoreEmpty Then
                ReDim Preserve arr(UBound(arr) - 1) ' when there is an empty cell, just reduce array size by 1
            Else
                arr(counter) = Values(I, J)
                counter = counter + 1
            End If
        Next J
    Next I
    
    RangeTo1DArray = arr
    
End Function

Private Function TestRangeTo2DArray()
    Debug.Print RangeTo2DArray(Range("A1"))(1, 1) ' should not throw an error
    Debug.Print RangeTo2DArray(Range("A1:B1"))(1, 2) ' should not throw an error
    Debug.Print RangeTo2DArray(Range("A1:A2"))(2, 1) ' should not throw an error
    Debug.Print RangeTo2DArray(Range("A1:B2"))(2, 2) ' should not throw an error
End Function

Public Function RangeTo2DArray(r As Range) As Variant
    ' ensure a range is converted to a 2-dimensional array
    ' special treatment on edge cases where a range is a 1x1 scalar
    If r.Cells.Count = 1 Then
        Dim arr() As Variant
        ReDim arr(1 To 1, 1 To 1) ' make it base 1, similar to what .value does for non-scalars
        arr(1, 1) = r.Value
        RangeTo2DArray = arr
    Else
        RangeTo2DArray = r.Value
    End If
    
End Function

'************"MiscRemoveGridLines"
'@IgnoreModule ImplicitByRefModifier


Public Sub RemoveGridLines(WS As Worksheet)
    Dim view As WorksheetView
    For Each view In WS.Parent.Windows(1).SheetViews
        If view.Sheet.Name = WS.Name Then
            view.DisplayGridlines = False
            Exit Sub
        End If
    Next
End Sub

'************"MiscString"
'@IgnoreModule ImplicitByRefModifier


Public Function randomString(length As Variant)
    Dim s As String
    While Len(s) < length
        s = s & Hex(Rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function


'************"MiscTables"
'@IgnoreModule ImplicitByRefModifier


Public Function HasLO(Name As String, Optional WB As Workbook) As Boolean

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If LCase(Name) = LCase(LO.Name) Then
                HasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    HasLO = False

End Function


' get list object only using it's name from within a workbook
Public Function GetLO(Name As String, Optional WB As Workbook) As ListObject

    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As Worksheet
    Dim LO As ListObject
    
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            If LCase(Name) = LCase(LO.Name) Then
                Set GetLO = LO
                Exit Function
            End If
        Next LO
    Next WS
    
    If GetLO Is Nothing Then
        ' 9: Subscript out of range
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "List object '" & Name & "' not found in workbook '" & WB.Name & "'")
    End If

End Function


Private Sub TestTableToArray()
    TableToArray "foo"
End Sub

Function TableToArray( _
      Name As String _
    , Optional WB As Workbook _
    ) As Variant()
    
    TableToArray = RangeTo2DArray(TableRange(Name, WB))
    
End Function

Public Function TableRange( _
        Name As String _
      , Optional WB As Workbook _
      ) As Range
    
'Returns the range (including headers of a table named `Name` in workbook `WB`): _
- It first looks for a list object called `Name` _
  - If the `.DataBodyRange` property is nothing the table range will only be the headers _
- Then it looks for a named range in the Workbook scope called `Name` and returns the _
  range this named range is referring to _
- Then it looks for a worksheet scoped named range called `Name`. The first occurrence _
  will be returned _
If no tables found, a `SubscriptOutOfRange` error (9) is raised _
The name of the table to be found is case insensitive
  
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    If HasLO(Name, WB) Then
        Dim LO As ListObject
        Set LO = GetLO(Name, WB)
        If LO.DataBodyRange Is Nothing Then
            Set TableRange = LO.HeaderRowRange
        Else
            Set TableRange = LO.Range
        End If
        Exit Function
    End If
    
    If hasKey(WB.Names, Name) Then
        Set TableRange = WB.Names(Name).RefersToRange
        Exit Function
    End If
    
    Dim WS As Worksheet
    ' this will find the first occurrence of the table called 'Name'
    For Each WS In WB.Worksheets
        If hasKey(WS.Names, Name) Then
            Set TableRange = WS.Names(Name).RefersToRange
            Exit Function
        End If
    Next WS
    
    Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Table '" & Name & "' not found in workbook '" & WB.Name & "'")
    
End Function


Function GetAllTables(WB As Workbook) As Collection
    Set GetAllTables = New Collection
    
    ' Returns all (potential) tables in a workbook
    
    Dim WS As Worksheet
    Dim LO As ListObject
    For Each WS In WB.Worksheets
        For Each LO In WS.ListObjects
            GetAllTables.Add LO.Name
        Next LO
    Next WS
    
    Dim Name As Name
    For Each Name In WB.Names
        GetAllTables.Add Name.Name
    Next Name
    
    For Each WS In WB.Worksheets
        For Each Name In WS.Names
            ' remove the sheetname prefix to get the table name
            GetAllTables.Add Mid(Name.Name, InStr(Name.Name, "!") + 1)
        Next Name
    Next WS
    
End Function


Function TableColumnToArray(TableDicts As Collection, ColumnName As String) As Variant()
    ' Converts a table's column to a 1-dimensional array
    
    Dim arr() As Variant
    ReDim arr(TableDicts.Count - 1) ' zero indexed
    Dim dict As Dictionary
    Dim counter As Long
    For Each dict In TableDicts
        arr(counter) = dictget(dict, ColumnName)
        counter = counter + 1 ' zero indexing
    Next dict
    
    TableColumnToArray = arr
End Function

'************"MiscTableToDicts"
'@IgnoreModule ImplicitByRefModifier


Private Sub TableToDictsTest()
    Dim Dicts As Collection
    Set Dicts = TableToDicts("TableToDictsTestData")
    ' read row 2 in column "b":
    Debug.Print Dicts(2)("b"), 5
End Sub

Public Function TableToDictsLogSource( _
          TableName As String _
        , Optional WB As Workbook _
        , Optional Columns As Collection _
        ) As Collection
    
'Similar to TableToDicts, but also stores the source of each row _
in a dictionary with key `__source__`

'The `__source__` object contains the following keys: _
 - `Workbook`: the Workbook object with the table _
 - `Table`: the name of the table within the workbook _
 - `RowIndex`: the row index of the current entry of the table

    Set TableToDictsLogSource = TableToDicts(TableName, WB, Columns)
    Dim dict As Dictionary
    Dim RowIndex As Long
    RowIndex = 0
    For Each dict In TableToDictsLogSource
        RowIndex = RowIndex + 1
        dict.Add "__source__", dicti("Workbook", WB, "Table", TableName, "RowIndex", RowIndex)
    Next dict
End Function


Public Function TableToDicts( _
          TableName As String _
        , Optional WB As Workbook _
        , Optional Columns As Collection _
        ) As Collection
    
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/8e1b9709a380d71eaf0d59bd0c2882c8501e9540/aa_py_xl/data_util.py#L21
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set TableToDicts = New Collection
    
    Dim d As Dictionary
    
    Dim I As Long
    Dim J As Long
    Dim TableData() As Variant
    TableData = TableToArray(TableName, WB)
    
    For I = LBound(TableData, 1) + 1 To UBound(TableData, 1)
        Set d = New Dictionary
        d.CompareMode = TextCompare ' must be case insensitive
        
        If Columns Is Nothing Then
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                d.Add TableData(1, J), TableData(I, J)
            Next J
        Else
            Dim ColumnName As Variant
            Dim Column As Variant
            
            For J = LBound(TableData, 2) To UBound(TableData, 2)
                ColumnName = TableData(LBound(TableData, 2), J)
                If IsValueInCollection(Columns, ColumnName) Then
                    d.Add ColumnName, TableData(I, J)
                End If
            Next J
        End If
        
        TableToDicts.Add d
    Next I
    
End Function

Private Function TestGetTableRowIndex()
    Dim Table As Collection
    Set Table = col(dicti("a", 1, "b", 2), dicti("a", 3, "b", 4), dicti("a", "foo", "b", "bar"))
    Debug.Print GetTableRowIndex(Table, col("a", "b"), col(3, 4)), 2
    Debug.Print GetTableRowIndex(Table, col("a", "b"), col("foo", "bar")), 3
End Function


Function TableLookupValue( _
        Table As Variant _
      , Columns As Collection _
      , Values As Collection _
      , ValueColName As String _
      , Optional default As Variant = Empty _
      , Optional WB As Workbook _
      ) As Variant
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' Returns the value from the ValueColName column in a TableToDicts object _
      given the value In the lookup column _
      A default value can be assigned For when no lookup Is found _
      Otherwise it returns a runtime Error
    
    ' for when GetTableRowIndex fails
    If Not IsEmpty(default) Then On Error GoTo SetDefault
    
    Dim dict As Dictionary
    Set dict = EnsureTableDicts(Table, WB)(GetTableRowIndex(Table, Columns, Values, WB))
    TableLookupValue = dictget(dict, ValueColName, default)
    
    Exit Function
SetDefault:
    TableLookupValue = default
    
End Function

Function GetTableRowRange( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Range
    
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, a runtime error is raised
    
    Dim RowNumber As Long
    RowNumber = GetTableRowIndex(TableName, Columns, Values, WB) ' this will throw a runtime error if not found
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    ' Intersect of table range and entirerow
    ' +1 as header is not included in GetTableRowIndex
    Set GetTableRowRange = Intersect(TableR, TableR(RowNumber + 1, 1).EntireRow)
    
End Function


Function GetTableColumnRange( _
      TableName As String _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    
' Returns the range of a table's column, including the header
    
    Dim TableR As Range
    Set TableR = TableRange(TableName, WB)
    
    Dim I As Long
    For I = 1 To TableR.Columns.Count
        If LCase(TableR(1, I).Value) = LCase(Column) Then
            GoTo found
        End If
    Next I
    
    Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Column '" & Column & "' not found in table '" & TableName & "'")
found:
    ' Intersect of table range and entirecolumn
    Set GetTableColumnRange = Intersect(TableR, TableR(1, I).EntireColumn)

End Function


Public Function TableLookupCell( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Column As String _
    , Optional WB As Workbook _
    ) As Range
    
    Set TableLookupCell = Intersect(GetTableRowRange(TableName, Columns, Values, WB), GetTableColumnRange(TableName, Column, WB))

End Function

Private Function EnsureTableDicts(Table As Variant, Optional WB As Workbook) As Collection
    
    If TypeOf Table Is Collection Then ' assume if collection, it's already a TableDicts object
        Set EnsureTableDicts = Table
    Else
        Set EnsureTableDicts = TableToDicts(CStr(Table), WB)
    End If

End Function


Function GetTableRowIndex( _
      Table As Variant _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    ) As Long
    
    ' Table can either be a TableToDicts collection, _
      or the name of the table to find
    
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, SubscriptOutOfRange error is raised
    
    Dim dict As Dictionary
    Dim keyValuePair As Collection
    Dim isMatch As Boolean
    Dim RowNumber As Long
    
    For Each dict In EnsureTableDicts(Table, WB)
        isMatch = True
        RowNumber = RowNumber + 1
        For Each keyValuePair In zip(Columns, Values)
            If dict(keyValuePair(1)) <> keyValuePair(2) Then
                isMatch = False
            End If
        Next keyValuePair
        If isMatch = True Then Exit For
    Next dict
    
    If isMatch Then
        GetTableRowIndex = RowNumber
    Else
        Err.Raise ErrNr.SubscriptOutOfRange, , ErrorMessage(ErrNr.SubscriptOutOfRange, "Columns-values pairs did not find a match")
    End If
    
End Function

Public Sub GotoRowInTable( _
      TableName As String _
    , Columns As Collection _
    , Values As Collection _
    , Optional WB As Workbook _
    )
    Application.GoTo GetTableRowRange(TableName, Columns, Values, WB), True
End Sub
