Attribute VB_Name = "MiscHasKey"
Option Explicit

Public Function hasKey(Container, key As Variant) As Boolean
    hasKey = True
    If TypeOf Container Is Collection Then
        On Error GoTo noKey
        TypeName Container(key)
        Exit Function
noKey:
        hasKey = False
    Else
        'We expect keyable VBA objects to have .Exists methods
        hasKey = Container.Exists(key)
        Exit Function
    End If
End Function
