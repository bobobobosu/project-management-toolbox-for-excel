Attribute VB_Name = "Module16"
Public Function IFSmaller(ori As Variant, compare As Variant, result As Variant) As Variant
    On Error Resume Next
    If ori < compare Then
        IFSmaller = result
    Else
        IFSmaller = ori
    End If
End Function

Public Function IFBigger(ori As Variant, compare As Variant, result As Variant) As Variant
    On Error Resume Next
    If ori > compare Then
        IFBigger = result
    Else
        IFBigger = ori
    End If
End Function

Public Function BandLimit(ori As Variant, Low As Variant, High As Variant) As Variant
    On Error Resume Next
    If ori > High Then
        BandLimit = High
    ElseIf ori < Low Then
        BandLimit = Low
    Else
        BandLimit = ori
    End If
End Function

Public Function ZeroAsNA(ori As Variant) As Variant
    If ori = 0 Then
        ZeroAsNA = CVErr(xlErrNA)
    Else
        ZeroAsNA = ori
    End If
End Function
