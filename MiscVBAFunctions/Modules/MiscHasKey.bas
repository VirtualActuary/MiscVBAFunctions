Attribute VB_Name = "MiscHasKey"
Option Explicit

Private Sub TestHasKey()

    Dim c As New Collection
    c.Add "a", "a"
    c.Add col("x", "y", "z"), "b"
    
    Debug.Print vbLf & "*********** TestHasKey tests ***********"
    Debug.Print hasKey(c, "a") ' True for scalar
    Debug.Print hasKey(c, "b") ' True for object
    Debug.Print hasKey(c, "A") ' False (case insensitive)

    Debug.Print hasKey(Workbooks, ThisWorkbook.Name) ' True for non-collection type collections
    
    Dim d As New Dictionary
    d.Add "a", "a"
    d.Add "b", col("x", "y", "z")
    
    Debug.Print hasKey(d, "a") ' True for scalar
    Debug.Print hasKey(d, "b") ' True for object
    Debug.Print hasKey(d, "A") ' False - case sensitive by default
    
    Dim dObj As Object
    Set dObj = CreateObject("Scripting.Dictionary")
    
    dObj.Add "a", "a"
    dObj.Add "b", col("x", "y", "z")
    
    Debug.Print hasKey(dObj, "a") ' True for scalar
    Debug.Print hasKey(dObj, "b") ' True for object
    Debug.Print hasKey(dObj, "A") ' False - case sensitive by default

End Sub

Public Function hasKey(Container, key As Variant) As Boolean
    hasKey = True
    If Not TypeOf Container Is Dictionary Then
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
