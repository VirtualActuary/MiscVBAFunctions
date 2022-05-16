Attribute VB_Name = "MiscErrorMessage"
Option Explicit
' #####################################################################
' ##### This module is version controlled in RunnerModule/Modules #####
' #####################################################################

'@Folder("error handling")

Public Function ErrorMessage(ErrorCode As Integer, _
                              Optional SubMessage As String) As String
    ' Get the Error message for the given error code and return it. If the ErrorCode is not
    ' in the list of known error codes, "Unknown error" will be the Error message.
    ' SubMessage can be added that will be appended to the String that will be returned.
    '
    ' Args:
    '   ErrorCode: Error code to look for.
    '   SubMessage: Message to append to the returned error message.
    '
    ' Returns:
    '   String with the error message and the SubMessage.
    
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

Public Function BreakLines(SubMessage As String) As String
    ' Split a word with 100+ characters into 3 lines and a word with 50+ words into 2 lines.
    '
    ' Args:
    '   SubMessage: String containing 1 or more words.
    '
    ' Returns:
    '   String with 50+ character words split between multiple lines.
    
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
