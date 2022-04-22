Attribute VB_Name = "MiscEarlyBindings"
'@IgnoreModule ImplicitByRefModifier

Option Explicit

' Add references for this project programatically. If you are uncertain what to put here,
' Go to `Tools -> References` and add the relevant bindings, then use the Sub
' printAllEarlyBindings to see how to add it as VBA code
'**********************************************************************************
'* Add selected references to this project
'**********************************************************************************
Sub addEarlyBindings()
    On Error GoTo ErrorHandler
    
        If Not isBindingNameLoaded("ADODB") Then
            'Microsoft ActiveX Data Objects 6.0 Library
            ThisWorkbook.VBProject.References.addFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6, 0
        End If
        

        If Not isBindingNameLoaded("VBIDE") Then
            'Microsoft Visual Basic for Applications Extensibility 5.3
            ThisWorkbook.VBProject.References.addFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
        End If


        If Not isBindingNameLoaded("Scripting") Then
            'Microsoft Scripting Runtime
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
    Dim codeString As String
    Dim codeStringTmp As String
    
    codeString = "" & _
        "If Not isBindingNameLoaded(""__name__"") Then" & chr(10) & _
        "    '__description__" & chr(10) & _
        "    ThisWorkbook.VBProject.References.addFromGuid ""__guid__"", __major__, __minor__" & chr(10) & _
        "End If" & chr(10)
        
        
    Dim xRef As Variant
    For Each xRef In ThisWorkbook.VBProject.References
        codeStringTmp = codeString
        codeStringTmp = Replace(codeStringTmp, "__name__", xRef.Name)
        codeStringTmp = Replace(codeStringTmp, "__description__", xRef.Description)
        codeStringTmp = Replace(codeStringTmp, "__guid__", xRef.GUID)
        codeStringTmp = Replace(codeStringTmp, "__major__", xRef.Major)
        codeStringTmp = Replace(codeStringTmp, "__minor__", xRef.Minor)
        
        Debug.Print
        Debug.Print codeStringTmp
        
    Next xRef
    
End Sub

