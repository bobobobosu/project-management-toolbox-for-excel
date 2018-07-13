Attribute VB_Name = "Module18"
Sub SetZero()
    For Each cell In Selection
        cell.Value2 = 0
    Next cell
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        Range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
