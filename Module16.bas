Attribute VB_Name = "Module16"
Public Function IFSmaller(ori As Variant, compare As Variant, result As Variant) As Variant
    If ori < compare Then
        IFSmaller = result
    Else
        IFSmaller = ori
    End If
End Function

Public Function IFBigger(ori As Variant, compare As Variant, result As Variant) As Variant
    If ori > compare Then
        IFBigger = result
    Else
        IFBigger = ori
    End If
End Function
