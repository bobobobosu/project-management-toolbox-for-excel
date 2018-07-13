Attribute VB_Name = "Module31"
Public Function ARRAYFIND(target1 As Variant, target2 As Range)

For Each cell In target2

If cell.Value = target1 Then
    variable1 = cell.Row
    variable2 = cell.Column
    variable3 = Cells(variable1, variable2).Address
    
    ARRAYFIND = variable3
    Exit Function
End If

Next


End Function


Public Function ARRAYINDEX(num As Integer, target2 As Range)
Dim count As Integer
Dim target As String
count = 0

For Each cell In target2

If cell.Value <> "" Then
    count = count + 1
End If


If count = num Then
    target = cell.Value
End If
If count = num Then Exit For

Next

ARRAYINDEX = target

End Function

