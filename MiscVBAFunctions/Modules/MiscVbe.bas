Attribute VB_Name = "MiscVbe"
' Utilities for manipulating the VBA Editor (VBE).
' See https://learn.microsoft.com/en-us/office/vba/api/excel.application.vbe

Option Explicit


Public Sub Compile()
    ' Compile the project. This is akin to clicking "Debug -> Compile VBAProject" in the menu.
    ' See https://stackoverflow.com/a/55613985/836995

    If IsCompiled() Then
        Exit Sub
    End If
    
    GetCompileCommand().Execute
End Sub


Public Function IsCompiled() As Boolean
    ' Check whether the current project is compiled.
    '
    ' As far as we can tell, this corresponds to whether the "Debug -> Compile VBAProject"
    ' menu option is disabled.
    '
    ' Returns:
    '   True if the project seems to be compiled.
    
    IsCompiled = Not GetCompileCommand().Enabled
End Function


Private Function GetCompileCommand() As CommandBarControl
    ' See https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.findcontrols
    Set GetCompileCommand = GetCommandBars.FindControl(Type:=MsoControlButton, ID:=578)
End Function


Private Function GetCommandBars() As CommandBars
    ' See https://learn.microsoft.com/en-us/office/vba/api/office.commandbars
    Set GetCommandBars = Application.VBE.CommandBars
End Function
