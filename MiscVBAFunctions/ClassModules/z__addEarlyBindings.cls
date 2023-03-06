VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "Z__addEarlyBindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Add references for this project programatically. If you are uncertain what to put here,
' Go to Tools -> References and use the filename of the reference (eg. msado15.dll for
' Microsoft ActiveX Data Objects 6.1 Library'), then run getPackageGUID("msado15.dll")
' to see what options you have:
'**********************************************************************************
'* Add selected references to this project
'**********************************************************************************
Public Sub AddEarlyBindings()
    On Error GoTo ErrorHandler
        If Not IsBindingNameLoaded("ADODB") Then
            'Microsoft ActiveX Data Objects 6.1
            ThisWorkbook.VBProject.References.AddFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6.1, 0
        End If
        
        If Not IsBindingNameLoaded("VBIDE") Then
            'Microsoft Visual Basic for Applications Extensibility 5.3
            ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5.3, 0
        End If
        
        
        If Not IsBindingNameLoaded("Scripting") Then
            'Microsoft Scripting Runtime version 1.0
            ThisWorkbook.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        End If
        
    
        If Not IsBindingNameLoaded("VBScript_RegExp_55") Then
            'Microsoft VBScript Regular Expressions 5.5
            ThisWorkbook.VBProject.References.AddFromGuid "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5
        End If
        
        If Not IsBindingNameLoaded("Shell32") Then
            'Microsoft Shell Controls And Automation
            ThisWorkbook.VBProject.References.AddFromGuid "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}", 1, 0
        End If
    
    Exit Sub
ErrorHandler:
End Sub


'**********************************************************************************
'* Verify if a reference is loaded
'**********************************************************************************
Public Function IsBindingNameLoaded(Ref As String) As Boolean
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    IsBindingNameLoaded = False
    Dim XRef As Variant
    For Each XRef In ThisWorkbook.VBProject.References
        If VBA.LCase(XRef.Name) = VBA.LCase(Ref) Then
            IsBindingNameLoaded = True
        End If
    Next XRef
    
End Function


'**********************************************************************************
'* Print all current active GUIDs
'**********************************************************************************
Public Sub PrintAllEarlyBindings()
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    Dim XRef As Variant
    For Each XRef In ThisWorkbook.VBProject.References
        Debug.Print "**************" & XRef.Name
        Debug.Print XRef.Description
        Debug.Print XRef.Major
        Debug.Print XRef.Minor
        Debug.Print XRef.FullPath
        Debug.Print XRef.GUID
        Debug.Print VbNullString
    Next XRef
    
End Sub

